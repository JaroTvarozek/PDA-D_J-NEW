// ==UserScript==
// @name         PDA Suite (HF Slovakia)
// @namespace    http://tampermonkey.net/
// @version      1.18.3
// @description  Vsetky vylepsenia PDA v jednom skripte + panel na zapinanie a vypinanie jednotlivych modulov
// @author       Gabris, Tvarozek
// @updateURL    https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/pda-suite.user.js
// @downloadURL  https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/pda-suite.user.js
// @match        https://hf.simplifier.cloud/*
// @icon         https://www.google.com/s2/favicons?sz=64&domain=simplifier.cloud
// @run-at       document-start
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_registerMenuCommand
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @connect      172.16.77.134
// @connect      mixinggroup.sharepoint.com
// @connect      sharepoint.com
// @connect      sharepointonline.com
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// ==/UserScript==

/*
 * ============================================================================
 *  PDA Suite - jeden skript namiesto osmich samostatnych
 * ============================================================================
 *
 *  Ako to funguje:
 *    - kazde vylepsenie je samostatny MODUL (dole v sekcii MODULY)
 *    - ktore moduly bezia sa nastavuje v paneli: ozubene koliesko vpravo dole
 *      na stranke PDA, alebo cez ikonu Tampermonkey -> "Nastavenia PDA Suite"
 *    - nastavenia sa ukladaju lokalne v Tampermonkey (GM storage), takze
 *      prezijeu aktualizaciu skriptu a NIE su sucastou repozitara
 *
 *  Citlive udaje (hesla kolegov, adresa sluzby vykresov) sa zadavaju
 *  v tom paneli - zamerne nie su v kode, lebo repozitar je verejny.
 *
 *  Pridanie noveho vylepsenia = pridat funkciu + jeden riadok do zoznamu
 *  MODULES. Nic ine netreba.
 * ============================================================================
 */

(function () {
    'use strict';

    const LOG = '[PDA Suite]';
    const W = unsafeWindow;

    /* ========================================================================
     *  1. ULOZISKO NASTAVENI
     * ====================================================================== */

    const KEY_MODULES = 'pda_modules_v1';
    const KEY_USERS = 'pda_users_v1';
    const KEY_PDM = 'pda_pdm_v1';
    const KEY_EXCEL = 'pda_excel_v1';
    const KEY_GROUPS = 'pda_groups_v1';
    const KEY_BUTTONS = 'pda_buttons_v1';
    const KEY_ADMIN = 'pda_admin_v1';

    function loadJson(key, fallback) {
        try {
            const raw = GM_getValue(key, null);
            if (raw === null || raw === undefined) return fallback;
            return typeof raw === 'string' ? JSON.parse(raw) : raw;
        } catch (e) {
            console.warn(LOG, 'nepodarilo sa nacitat nastavenie', key, e);
            return fallback;
        }
    }

    function saveJson(key, value) {
        try {
            GM_setValue(key, JSON.stringify(value));
        } catch (e) {
            console.warn(LOG, 'nepodarilo sa ulozit nastavenie', key, e);
        }
    }

    const settings = {
        modules: loadJson(KEY_MODULES, {}),
        users: loadJson(KEY_USERS, []),
        pdm: loadJson(KEY_PDM, { base: 'http://172.16.77.134:9000', key: '' }),
        // url prazdna = subor sa vybera rucne cez tlacidlo "Vybrať Excel"
        excel: loadJson(KEY_EXCEL, { url: '', colOrder: 'H', colDrawing: 'AH', colVersion: 'AI' }),
        // kategorie pracovisk pre uvodny prehlad: Nazov|Podtitul|VZOR,VZOR,...
        groups: loadJson(KEY_GROUPS, [
            'Assembly|Finálna montáž a podzostavy|MONTAZ,ASSEMBLY,PODZOST,MONT',
            'Welding|Zváranie a príprava|ZVAR,TIG,MIG,WELD',
            'Machining|CNC a konvenčné obrábanie|CNC,FREZ,SUSTR,BRUS,LMS,HMS,K-TEC,VRTA,PILA,HOBL,HEDELL',
            'Quality Control|Kontrola a meranie|OTK,KONTROL,MERAN,QC,KVALIT',
        ].join('\n')),
        // pravidla farieb tlacidiel; null = pri prvom spusteni sa nasadia predvolene
        buttons: loadJson(KEY_BUTTONS, null),
        // heslo na ozubene koliesko a prepinac 'nastavovanie tlacidiel pravym klikom'
        admin: loadJson(KEY_ADMIN, { password: '123456', pickMode: false }),
    };

    // "H" -> 7, "AH" -> 33 (vracia 0-based index stlpca ako ho vidi XLSX)
    function colToIndex(letters, fallbackIndex) {
        const s = String(letters || '').trim().toUpperCase();
        if (!/^[A-Z]{1,3}$/.test(s)) return fallbackIndex;
        let n = 0;
        for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64);
        return n - 1;
    }

    function isModuleOn(mod) {
        const stored = settings.modules[mod.id];
        return stored === undefined ? mod.def : !!stored;
    }

    /* ========================================================================
     *  2. ZDIELANE POMOCKY
     * ====================================================================== */

    function sleep(ms) {
        return new Promise((r) => setTimeout(r, ms));
    }

    function onReady(fn) {
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', fn, { once: true });
        } else {
            fn();
        }
    }

    function waitFor(checkFn, { interval = 150, timeout = 10000 } = {}) {
        return new Promise((resolve, reject) => {
            const start = Date.now();
            const timer = setInterval(() => {
                let result = null;
                try { result = checkFn(); } catch (e) { /* ignore */ }
                if (result) {
                    clearInterval(timer);
                    resolve(result);
                } else if (Date.now() - start > timeout) {
                    clearInterval(timer);
                    reject(new Error('waitFor timeout'));
                }
            }, interval);
        });
    }

    // --- pristup k SAP UI5 controlom (aplikacia bezi v kontexte stranky) ---

    function getControl(controlId) {
        try {
            if (W.sap && W.sap.ui && W.sap.ui.getCore) return W.sap.ui.getCore().byId(controlId);
        } catch (e) { /* ignore */ }
        return null;
    }

    function resolveControl(el) {
        while (el && el.nodeType === 1) {
            try {
                if (W.jQuery && typeof W.jQuery(el).control === 'function') {
                    const arr = W.jQuery(el).control();
                    if (arr && arr.length && arr[0]) return arr[0];
                }
                if (W.sap && W.sap.ui && W.sap.ui.core && W.sap.ui.core.Element &&
                    typeof W.sap.ui.core.Element.closestTo === 'function') {
                    const c = W.sap.ui.core.Element.closestTo(el);
                    if (c) return c;
                }
                if (el.id && W.sap && W.sap.ui && W.sap.ui.getCore) {
                    const c = W.sap.ui.getCore().byId(el.id);
                    if (c) return c;
                }
            } catch (e) { /* ignore */ }
            el = el.parentElement;
        }
        return null;
    }

    function pressElement(domEl) {
        const control = resolveControl(domEl);
        if (control && typeof control.firePress === 'function') {
            control.firePress();
            return;
        }
        domEl.click();
    }

    function setControlValue(controlId, value) {
        try {
            const control = getControl(controlId);
            if (control && typeof control.setValue === 'function') {
                control.setValue(value);
                if (typeof control.fireLiveChange === 'function') control.fireLiveChange({ value });
                if (typeof control.fireChange === 'function') control.fireChange({ value, newValue: value });
                return true;
            }
        } catch (e) {
            console.warn(LOG, 'setControlValue zlyhalo pre', controlId, e);
        }
        const inputEl = document.getElementById(controlId + '-inner');
        if (inputEl) {
            const nativeSetter = Object.getOwnPropertyDescriptor(W.HTMLInputElement.prototype, 'value').set;
            nativeSetter.call(inputEl, value);
            inputEl.dispatchEvent(new Event('input', { bubbles: true }));
            inputEl.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        return false;
    }

    /* ------------------------------------------------------------------
     *  Jeden spolocny sledovac DOM namiesto piatich samostatnych.
     *  Aplikacia je SAP UI5 a prekresluje sa sama, takze moduly musia
     *  svoje prvky dokladat opakovane. Vsetko sa zbiera do jednej davky.
     * ---------------------------------------------------------------- */
    const DomWatch = (function () {
        const callbacks = [];
        let started = false;
        let scheduled = false;

        function flush() {
            scheduled = false;
            for (const fn of callbacks) {
                try { fn(); } catch (e) { console.warn(LOG, 'chyba v DOM callbacku', e); }
            }
        }

        function schedule() {
            if (scheduled) return;
            scheduled = true;
            setTimeout(flush, 120);
        }

        function start() {
            if (started || !document.body) return;
            started = true;
            new MutationObserver(schedule).observe(document.body, {
                childList: true, subtree: true, characterData: true,
            });
        }

        return {
            add(fn) {
                callbacks.push(fn);
                onReady(() => { start(); schedule(); });
            },
            poke: schedule,
        };
    })();

    /* ------------------------------------------------------------------
     *  Jedno spolocne odpocuvanie sietovej komunikacie namiesto troch.
     *  Prepisuje XMLHttpRequest v kontexte stranky a rozposiela vysledky
     *  vsetkym prihlasenym modulom.
     * ---------------------------------------------------------------- */
    const XhrBus = (function () {
        const TARGET = '/client/1.0/executeBO';
        const listeners = [];
        let patched = false;

        function patch() {
            if (patched) return;
            patched = true;

            const proto = W.XMLHttpRequest.prototype;
            const originalOpen = proto.open;
            const originalSend = proto.send;

            proto.open = function (method, url, ...rest) {
                this.__pdaUrl = url;
                this.__pdaMethod = method;
                return originalOpen.call(this, method, url, ...rest);
            };

            proto.send = function (body) {
                const url = this.__pdaUrl || '';
                if (url.indexOf(TARGET) !== -1) {
                    const xhr = this;
                    xhr.addEventListener('load', function () {
                        let requestJson = null;
                        let responseJson = null;
                        try { requestJson = JSON.parse(body); } catch (e) { /* ignore */ }
                        try { responseJson = JSON.parse(xhr.responseText); } catch (e) { /* ignore */ }

                        const event = {
                            url,
                            requestRaw: body,
                            request: requestJson,
                            response: responseJson,
                            responseRaw: xhr.responseText,
                        };
                        for (const fn of listeners) {
                            try { fn(event); } catch (e) { console.warn(LOG, 'chyba v XHR callbacku', e); }
                        }
                    });
                }
                return originalSend.apply(this, arguments);
            };
        }

        return {
            subscribe(fn) {
                patch();
                listeners.push(fn);
            },
        };
    })();

    /* ------------------------------------------------------------------
     *  Dekodovanie vstupu zo skenera.
     *  Skener posiela cisla ako slovenske znaky (SK klavesnica), prvy znak
     *  "J" je marker zariadenia a bodka zastupuje pomlcku.
     * ---------------------------------------------------------------- */
    const SCANNER_CHAR_MAP = {
        '+': '1', 'ľ': '2', 'š': '3', 'č': '4', 'ť': '5',
        'ž': '6', 'ý': '7', 'á': '8', 'í': '9', 'é': '0',
    };

    function decodeScannerInput(raw) {
        // POZOR: prvy znak sa odstrani LEN ak je to naozaj marker "J".
        // (v povodnych skriptoch sa na jednom mieste odrezaval vzdy, co
        //  rucne napisanemu cislu zjedlo prvu cislicu)
        const stripped = raw.length > 0 && raw[0] === 'J' ? raw.slice(1) : raw;

        let out = '';
        for (const ch of stripped) {
            if (ch === '.') out += '-';
            else if (Object.prototype.hasOwnProperty.call(SCANNER_CHAR_MAP, ch)) out += SCANNER_CHAR_MAP[ch];
            else out += ch;
        }
        return out;
    }

    function padOperationPart(value) {
        const i = value.indexOf('-');
        if (i === -1) return value;
        return value.slice(0, i) + '-' + value.slice(i + 1).padStart(4, '0');
    }

    /* ------------------------------------------------------------------
     *  Prepinanie pouzivatela - zdielana sluzba.
     *  Pouziva ju aj bocny panel, aj citacka kariet, takze je tu zvlast
     *  a nie je viazana na to, ci je panel zapnuty.
     * ---------------------------------------------------------------- */
    const UserSwitch = (function () {
        const CHANGE_USER_SUFFIX = 'Button_ChangeUser';
        const USER_COMBO_ID = 'Popups--User_ComboBox';
        const PASSWORD_INPUT_ID = 'Popups--EmployeeNo_Input';
        const CONFIRM_BUTTON_ID = 'Popups--UserDialog_Button_close';
        const DIALOG_WAIT_TIMEOUT = 10000;
        const BEFORE_CONFIRM_DELAY = 100;

        async function selectComboBoxItemByText(controlId, text, { timeout = 5000 } = {}) {
            let control;
            try {
                control = await waitFor(() => {
                    const c = getControl(controlId);
                    return c && typeof c.getItems === 'function' && c.getItems().length > 0 ? c : null;
                }, { timeout });
            } catch (e) {
                console.warn(LOG, 'polozky v zozname pouzivatelov sa nenacitali vcas', e);
                return false;
            }

            const items = control.getItems();
            const match =
                items.find((it) => it.getText() === text) ||
                items.find((it) => it.getText().toLowerCase() === text.toLowerCase());

            if (!match) {
                console.warn(LOG, 'pouzivatel sa v zozname nenasiel:', text,
                    'dostupni:', items.map((i) => i.getText()));
                return false;
            }

            control.setSelectedItem(match);
            if (typeof control.fireSelectionChange === 'function') control.fireSelectionChange({ selectedItem: match });
            if (typeof control.fireChange === 'function') control.fireChange({ value: match.getText() });
            return true;
        }

        async function fillLoginPopup(username, password) {
            try {
                await waitFor(
                    () => document.getElementById(USER_COMBO_ID + '-inner') &&
                          document.getElementById(PASSWORD_INPUT_ID + '-inner'),
                    { timeout: DIALOG_WAIT_TIMEOUT }
                );

                const selected = await selectComboBoxItemByText(USER_COMBO_ID, username);
                if (!selected) setControlValue(USER_COMBO_ID, username);

                setControlValue(PASSWORD_INPUT_ID, password);
                await sleep(BEFORE_CONFIRM_DELAY);

                const confirmBtn = await waitFor(() => document.getElementById(CONFIRM_BUTTON_ID), { timeout: 5000 });
                pressElement(confirmBtn);
            } catch (e) {
                console.warn(LOG, 'dialog na zmenu pouzivatela sa neobjavil vcas', e);
            }
        }

        return {
            to(username, password) {
                const btn = document.querySelector('[id$="' + CHANGE_USER_SUFFIX + '"]');
                if (!btn) {
                    console.warn(LOG, 'tlacidlo na zmenu pouzivatela sa na obrazovke nenaslo');
                    return;
                }
                pressElement(btn);
                fillLoginPopup(username, password);
            },
            findByCardId(cardId) {
                return settings.users.find((u) => u.cardId && u.cardId === cardId) || null;
            },
        };
    })();

    // zdielany stav medzi modulmi (nahradza povodne window.PDA_* premenne)
    const shared = {
        ordersIndex: [],
        ordersIndexUpdatedAt: null,
        currentOperation: null,
        fillSearchInput: null,   // doplni modul vyhladavania, ak bezi
        intentionalReload: false, // nastavi panel nastaveni pred location.reload()
    };

    /* ========================================================================
     *  3. MODULY
     * ====================================================================== */

    /* ---------------------- 3.1 Vylepsena hlavicka ---------------------- */

    function modEnhancedHeader() {
        const USERNAME_SUFFIX = 'Label_Username-bdi';
        const BUTTONS = [
            { suffix: 'Button_HomeScreen-img', label: 'Pracoviská' },
            { suffix: 'Button_Reporting-img', label: 'Reporty' },
            { suffix: 'Button_Message-img', label: 'Správy' },
            { suffix: 'Button_Schedule-img', label: 'Rozvrh' },
            { suffix: 'Button_Settings-img', label: 'Admin' },
            { suffix: 'Button_ChangeUser-img', label: 'Zmena používateľa' },
            { suffix: 'Button_Logout-img', label: 'Odhlásenie' },
        ];
        const LOGOUT_BUTTON_SUFFIX = 'Button_Logout-inner';

        function apply() {
            document.querySelectorAll('[id$="' + USERNAME_SUFFIX + '"]').forEach((el) => {
                if (el.dataset.pdaUsernameStyled) return;
                el.style.fontSize = '1.8rem';
                el.style.fontWeight = 'bold';
                el.style.color = '#1f1f1f';
                el.dataset.pdaUsernameStyled = '1';
            });

            BUTTONS.forEach(({ suffix, label }) => {
                document.querySelectorAll('[id$="' + suffix + '"]').forEach((img) => {
                    if (img.dataset.pdaTextAdded) return;
                    const span = document.createElement('span');
                    span.className = 'sapMBtnContent';
                    span.style.marginRight = '0.6rem';
                    span.innerHTML = '<bdi>' + label + '</bdi>';
                    img.insertAdjacentElement('afterend', span);
                    img.dataset.pdaTextAdded = '1';
                });
            });

            document.querySelectorAll('[id$="' + LOGOUT_BUTTON_SUFFIX + '"]').forEach((btn) => {
                if (btn.dataset.pdaLogoutStyled) return;
                btn.style.backgroundColor = '#d67a74';
                btn.style.borderRadius = '4px';
                btn.dataset.pdaLogoutStyled = '1';
            });
        }

        DomWatch.add(apply);
    }

    /* ------------------ 3.2 Farebne tlacidla (pravidla) ----------------- */

    /*
     * Portovane z Python appky (SKIN + PICKER v pda_action.py, karta "Farby tlacidiel").
     *  - pravidla { text, id, bg, fg, poradie } su v nastaveniach (ulozisko Tampermonkey),
     *    NIE v kode; pri prvom spusteni sa nasadia predvolene farby stavovych tlacidiel,
     *    takze sa oproti doterajsiemu spravaniu nic nemeni
     *  - matchuje sa podla TEXTU tlacidla ("obsahuje", bez ohladu na velkost pismen);
     *    ID sa pouzije len ked tlacidlo text nema - ID stavovych tlacidiel obsahuje
     *    poradove cislo, ktore sa lisi podla pracoviska
     *  - farba textu sa dopocita z jasu pozadia
     *  - pravy klik na tlacidlo (ked je v nastaveniach zapnute "Nastavovanie tlacidiel")
     *    ukaze paletu 12 farieb + Reset; vyber sa hned ulozi a nanesie
     *  - `poradie` drzi doterajsie zoradenie stavovych tlacidiel (vyroba, prestoj, chyba)
     */
    const DEFAULT_BUTTON_RULES = [
        { text: 'Výroba', id: '', bg: '#4e9041', fg: '#ffffff', poradie: 0 },
        { text: 'Upinanie', id: '', bg: '#4e9041', fg: '#ffffff', poradie: 0 },
        { text: 'Programovanie', id: '', bg: '#d66c37', fg: '#ffffff', poradie: 1 },
        { text: 'Upratovanie stola', id: '', bg: '#d66c37', fg: '#ffffff', poradie: 1 },
        { text: 'Meranie v Procese s OTK', id: '', bg: '#d66c37', fg: '#ffffff', poradie: 1 },
        { text: 'Chyba programu', id: '', bg: '#d04040', fg: '#ffffff', poradie: 2 },
    ];

    // rovnakych 12 farieb ako v Python verzii
    const BUTTON_PALETTE = ['#16a34a', '#65a30d', '#eab308', '#d97706', '#dc2626', '#db2777',
                            '#7c3aed', '#2563eb', '#0891b2', '#7c4a1e', '#64748b', '#1e293b'];

    function contrastColor(hex) {
        const h = String(hex || '').replace('#', '');
        if (h.length !== 6) return '#ffffff';
        const r = parseInt(h.slice(0, 2), 16), g = parseInt(h.slice(2, 4), 16), b = parseInt(h.slice(4, 6), 16);
        return (0.299 * r + 0.587 * g + 0.114 * b) > 150 ? '#1e293b' : '#ffffff';
    }

    function normalizeRule(r) {
        const bg = /^#[0-9a-f]{6}$/i.test(String(r.bg || '')) ? String(r.bg).toLowerCase() : '#2563eb';
        const fg = /^#[0-9a-f]{6}$/i.test(String(r.fg || '')) ? String(r.fg).toLowerCase() : contrastColor(bg);
        const poradie = Number.isFinite(Number(r.poradie)) ? Number(r.poradie) : 1.5;
        return { text: String(r.text || '').trim(), id: String(r.id || '').trim(), bg, fg, poradie };
    }

    function buttonRules() {
        if (!Array.isArray(settings.buttons)) {
            settings.buttons = DEFAULT_BUTTON_RULES.map((r) => Object.assign({}, r));
        }
        return settings.buttons;
    }

    function saveButtonRules(rules) {
        settings.buttons = rules.map(normalizeRule).filter((r) => r.text || r.id);
        saveJson(KEY_BUTTONS, settings.buttons);
        mirrorSettingsToFile(false);
    }

    // rovnake pravidlo ako Python `bezPravidla`: pravidlo "patri" tlacidlu podla textu alebo ID
    function ruleMatches(rule, txt, id) {
        const t = String(rule.text || '').toLowerCase();
        return (t && String(txt || '').toLowerCase().indexOf(t) !== -1) || (rule.id && rule.id === id);
    }

    function modButtonColors() {
        const CONTAINER_ID = 'WorkcenterDetail--Order_Status_Flexbox';
        const OWN_UI = '#__pda_settings_overlay__, #__pda_settings_pass__, #__pda_overview__, #__pda_pdm_overlay__, #__pda_button_menu__';
        const STYLE_ID = '__pda_status_buttons_styles__';

        /*
         * Tvar stavovych tlacidiel (farby nanasa `paint` inline podla pravidiel):
         * povodne boli vysoke cez 100 px a pri viacerych riadkoch bola medzi
         * riadkami velka diera. Teraz su nizsie, s jemnym ramom a tienom a pri
         * prechode mysou sa o 2 px nadvihnu.
         */
        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* vsetky tlacidla v riadku su rovnako vysoke - natiahnu sa na to najvyssie
   (dvojriadkove "Meranie v Procese s OTK"), text ostava zvisle na stred */
#${CONTAINER_ID} { align-content:flex-start !important; align-items:stretch !important; row-gap:0 !important; }
#${CONTAINER_ID} .statusBtn { height:auto !important; min-height:0 !important; margin:4px !important;
  align-self:stretch !important; border-radius:12px !important;
  border:2px solid #13315c !important;
  box-shadow:0 2px 6px rgba(16,36,63,.20) !important;
  transition:transform .13s ease, box-shadow .13s ease !important; }
#${CONTAINER_ID} .statusBtn .sapMBtnInner { height:100% !important; width:100% !important; min-height:0 !important;
  padding:11px 16px !important; border-radius:12px !important; box-shadow:none !important;
  display:flex !important; align-items:center !important; justify-content:center !important;
  box-sizing:border-box !important; }
#${CONTAINER_ID} .statusBtn .sapMBtnContent, #${CONTAINER_ID} .statusBtn bdi {
  line-height:1.25 !important; }
#${CONTAINER_ID} .statusBtn:hover { transform:translateY(-3px) !important;
  box-shadow:0 10px 20px rgba(16,36,63,.30) !important; }
#${CONTAINER_ID} .statusBtn:active { transform:translateY(-1px) !important;
  box-shadow:0 3px 8px rgba(16,36,63,.24) !important; }
`;
            document.head.appendChild(st);
        }

        function textOf(btn) {
            return (btn.textContent || '').trim();
        }

        // najkonkretnejsie (najdlhsie) textove pravidlo vyhrava; ID len ak nic textove nesedi
        function ruleFor(btn) {
            const txt = textOf(btn).toLowerCase();
            const id = btn.id || '';
            let best = null, byId = null;
            for (const r of buttonRules()) {
                const t = String(r.text || '').toLowerCase();
                if (t) {
                    if (txt.indexOf(t) !== -1 && (!best || t.length > String(best.text).length)) best = r;
                } else if (r.id && r.id === id) {
                    byId = r;
                }
            }
            return best || byId;
        }

        function paint(btn, rule) {
            const inner = btn.querySelector('.sapMBtnInner') || btn;
            const content = btn.querySelector('.sapMBtnContent');
            if (rule) {
                inner.style.setProperty('background-image', 'none', 'important'); // gradient temy by farbu prekryl
                inner.style.setProperty('background-color', rule.bg, 'important');
                inner.style.setProperty('border-color', rule.bg, 'important');
                inner.style.setProperty('color', rule.fg, 'important');
                if (content) content.style.setProperty('color', rule.fg, 'important');
                btn.dataset.pdaPainted = '1';
            } else if (btn.dataset.pdaPainted) {
                ['background-image', 'background-color', 'border-color', 'color'].forEach((p) => inner.style.removeProperty(p));
                if (content) content.style.removeProperty('color');
                delete btn.dataset.pdaPainted;
            }
        }

        function apply() {
            injectStyles();
            document.querySelectorAll('.sapMBtn').forEach((btn) => {
                if (btn.closest(OWN_UI)) return;
                paint(btn, ruleFor(btn));
                if (btn.classList.contains('statusBtn')) {
                    [btn, btn.querySelector('.sapMBtnInner')].forEach((el) => el && el.style.setProperty('border-radius', '10px', 'important'));
                }
            });

            // zoradenie stavovych tlacidiel podla `poradie` (0 vyroba, 1 prestoj, 2 chyba, inak 1.5)
            const container = document.getElementById(CONTAINER_ID);
            if (!container) return;
            const buttons = Array.from(container.children).filter((el) => el.classList.contains('statusBtn'));
            if (buttons.length < 2) return;
            const weight = (btn) => {
                const r = ruleFor(btn);
                return r && Number.isFinite(Number(r.poradie)) ? Number(r.poradie) : 1.5;
            };
            const sorted = [...buttons].sort((a, b) => weight(a) - weight(b));
            if (!sorted.every((b, i) => buttons[i] === b)) sorted.forEach((b) => container.appendChild(b));
        }

        DomWatch.add(apply);

        /* ---- pravy klik: paleta farieb (len ked je v nastaveniach zapnute) ---- */

        let menu = null;

        function closeMenu() {
            if (menu) { menu.remove(); menu = null; }
        }

        function openMenu(e, btn) {
            closeMenu();
            const txt = textOf(btn).slice(0, 60);
            const id = btn.id || '';

            menu = document.createElement('div');
            menu.id = '__pda_button_menu__';
            menu.style.cssText =
                'position:fixed;z-index:2147483002;background:#13315c;color:#fff;' +
                'font:12px/1.4 "Segoe UI",system-ui,sans-serif;border-radius:12px;padding:12px 14px;' +
                'box-shadow:0 6px 20px rgba(0,0,0,.35);max-width:440px;';
            menu.style.left = Math.max(4, Math.min(e.clientX, W.innerWidth - 460)) + 'px';
            menu.style.top = Math.max(4, Math.min(e.clientY, W.innerHeight - 170)) + 'px';

            const title = document.createElement('div');
            title.style.cssText = 'font-weight:700;margin-bottom:4px;';
            title.textContent = txt || '(bez textu)';
            const sub = document.createElement('div');
            sub.style.cssText = 'color:#a8c0e0;font-size:10px;word-break:break-all;margin-bottom:8px;';
            sub.textContent = id || '(bez ID)';
            const label = document.createElement('div');
            label.style.cssText = 'color:#a8c0e0;font-size:11px;margin-bottom:4px;';
            label.textContent = 'Zmeniť farbu';

            const swatches = document.createElement('div');
            swatches.style.marginBottom = '10px';
            BUTTON_PALETTE.forEach((c) => {
                const sw = document.createElement('span');
                sw.style.cssText = 'display:inline-block;width:28px;height:28px;border-radius:8px;margin:3px;cursor:pointer;' +
                    'border:1px solid rgba(255,255,255,.28);background:' + c + ';';
                sw.title = c;
                sw.addEventListener('click', (ev) => {
                    ev.stopPropagation();
                    const fg = contrastColor(c);
                    const old = buttonRules().find((r) => ruleMatches(r, txt, id));
                    const rest = buttonRules().filter((r) => !ruleMatches(r, txt, id));
                    rest.push({ text: txt, id: txt ? '' : id, bg: c, fg, poradie: old ? old.poradie : 1.5 });
                    saveButtonRules(rest);
                    paint(btn, { bg: c, fg });
                    DomWatch.poke();
                    closeMenu();
                });
                swatches.appendChild(sw);
            });

            const reset = document.createElement('button');
            reset.type = 'button';
            reset.textContent = 'Reset tlačidla';
            reset.style.cssText = 'background:transparent;color:#a8c0e0;border:1px solid #1c478a;border-radius:9px;' +
                'padding:7px 14px;cursor:pointer;font:inherit;font-size:12px;';
            reset.addEventListener('click', (ev) => {
                ev.stopPropagation();
                saveButtonRules(buttonRules().filter((r) => !ruleMatches(r, txt, id)));
                paint(btn, null);
                DomWatch.poke();
                closeMenu();
            });

            menu.appendChild(title);
            menu.appendChild(sub);
            menu.appendChild(label);
            menu.appendChild(swatches);
            menu.appendChild(reset);
            document.body.appendChild(menu);
        }

        document.addEventListener('click', (e) => { if (menu && !menu.contains(e.target)) closeMenu(); }, true);
        document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeMenu(); }, true);
        document.addEventListener('contextmenu', (e) => {
            if (!settings.admin || !settings.admin.pickMode) return;
            const btn = e.target && e.target.closest ? e.target.closest('.sapMBtn') : null;
            if (!btn || btn.closest(OWN_UI)) return;
            e.preventDefault();
            e.stopPropagation();
            openMenu(e, btn);
        }, true);
    }

    /* ------------------ 3.3 Blokovanie tlacidla Spat -------------------- */

    function modPreventBack() {
        function lockCurrentState() {
            W.history.pushState(W.history.state, document.title, location.href);
        }

        onReady(() => {
            lockCurrentState();
            W.addEventListener('popstate', lockCurrentState);
            W.addEventListener('beforeunload', (e) => {
                // obnovenie stranky vyvolane panelom nastaveni sa nema pytat
                if (shared.intentionalReload) return;
                e.preventDefault();
                e.returnValue = '';
            });
        });
    }

    /* --------------- 3.4 Vyhladavanie naprieč pracoviskami -------------- */

    function modCrossSearch() {
        const BO_METHOD = 'getOperationListTempForWorkcenters';
        const WORKCENTER_TILE_SELECTOR = '[id^="Main--Workcenter_Toolbar-Main--ui_layout_Grid3-"]';
        const CONTENT_ID = 'Main--Workcenter_Panel-content';
        const HOME_BUTTON_ID = 'Main--Button_HomeScreen';
        const PANEL_ID = 'Main--Workcenter_Panel';
        const PARENT_SECTION_ID = 'Main--MainPage-cont';
        const TASKS_PANEL_ID = 'Main--Tasks_Panel';
        const SIDEBAR_ID = '__pda_search_sidebar__';
        const UI_ID = '__pda_custom_search_ui__';
        const TILE_ID_PREFIX = 'Main--Workcenter_Toolbar-';
        const LIST_ID_PREFIX = 'Main--List2-';
        const TILE_WAIT_TIMEOUT = 15000;
        const SETTLE_DELAY = 350;

        let opening = false;
        let renderFn = null;
        let lastAutoOpenedKey = null;

        XhrBus.subscribe((ev) => {
            if (!ev.request || ev.request.BOMethod !== BO_METHOD) return;
            if (!ev.response || !ev.response.result) return;

            const workcenters = ev.response.result.aOperationListTempForWorkcenter || [];
            const flat = [];
            workcenters.forEach((wc) => {
                (wc.operationList || []).forEach((op) => {
                    flat.push({
                        workcenter: wc.workcenterDescription,
                        workcenterCode: wc.workcenter,
                        salesOrderNo: op.salesOrderNo,
                        salesOrderItem: op.salesOrderItem,
                        productionOrderNo: op.productionOrderNo,
                        operationNo: op.operationNo,
                        sequenceNo: op.sequenceNo,
                        materialNo: op.materialNo,
                        material: op.material,
                        description: op.description,
                        descriptionShort: op.descriptionShort,
                        status: op.status,
                        confirmed: op.confirmed,
                    });
                });
            });

            shared.ordersIndex = flat;
            shared.ordersIndexUpdatedAt = new Date();
            console.log(LOG, 'index zakaziek naplneny:', flat.length);

            if (renderFn) {
                const input = document.querySelector('#' + UI_ID + ' input.sapMSFI');
                renderFn(input ? input.value : '');
            }
        });

        function getWorkcenterName(tile) {
            const suffix = tile.id.replace(TILE_ID_PREFIX, '');
            const nameEl = document.getElementById('Main--WorkcenterDescription_Label-' + suffix + '-bdi');
            return nameEl ? nameEl.textContent.trim() : tile.id;
        }

        function getListIdForTile(tile) {
            if (!tile.id || tile.id.indexOf(TILE_ID_PREFIX) !== 0) return null;
            return LIST_ID_PREFIX + tile.id.slice(TILE_ID_PREFIX.length);
        }

        function extractOrderText(li) {
            const titleEl = li.querySelector('[id*="Title"][id$="-inner"]');
            return titleEl ? titleEl.textContent.trim() : '';
        }

        function formatProductionOrder(item) {
            return item.productionOrderNo + ' - ' + item.operationNo + ' - ' + item.sequenceNo;
        }

        function getItemKey(item) {
            return [item.workcenter, item.productionOrderNo, item.operationNo, item.sequenceNo].join('|');
        }

        function setStatus(msg) {
            const el = document.getElementById('__pda_status__');
            if (el) el.textContent = msg;
        }

        function fireListItemPress(li) {
            const itemControl = resolveControl(li);
            if (!itemControl) return false;

            let listControl = itemControl;
            while (listControl && typeof listControl.fireItemPress !== 'function') {
                listControl = listControl.getParent && listControl.getParent();
            }
            if (!listControl) return false;

            let itemForEvent = itemControl;
            while (itemForEvent && itemForEvent.getParent && itemForEvent.getParent() !== listControl) {
                itemForEvent = itemForEvent.getParent();
            }
            if (!itemForEvent) itemForEvent = itemControl;

            listControl.fireItemPress({ listItem: itemForEvent, srcControl: itemForEvent });
            return true;
        }

        async function openItem(item) {
            if (opening) return;
            opening = true;
            try {
                setStatus('Otváram zákazku...');

                const tiles = Array.from(document.querySelectorAll(WORKCENTER_TILE_SELECTOR));
                const tile = tiles.find((t) => getWorkcenterName(t) === item.workcenter);
                if (!tile) { setStatus('Pracovisko sa nenašlo.'); return; }

                const listId = getListIdForTile(tile);
                let list = listId ? document.getElementById(listId) : null;

                if (!list) {
                    pressElement(tile);
                    list = listId ? await waitFor(() => document.getElementById(listId), { timeout: TILE_WAIT_TIMEOUT }) : null;
                }
                if (!list) { setStatus('Zoznam zákaziek pre toto pracovisko sa nenašiel.'); return; }

                list.scrollIntoView({ block: 'center' });
                await sleep(SETTLE_DELAY);

                const targetText = formatProductionOrder(item);
                let li = Array.from(list.querySelectorAll('li')).find((el) => extractOrderText(el) === targetText);
                if (!li) {
                    await sleep(SETTLE_DELAY);
                    li = Array.from(list.querySelectorAll('li')).find((el) => extractOrderText(el) === targetText);
                }
                if (!li) { setStatus('Konkrétna zákazka sa nenašla, otvorené je aspoň pracovisko.'); return; }

                li.scrollIntoView({ block: 'center' });
                const fired = fireListItemPress(li);
                if (!fired) li.click();
                setStatus(fired ? 'Otvorené: ' + targetText : 'Chyba pri otváraní zákazky.');
            } catch (e) {
                console.warn(LOG, 'chyba pri otvarani polozky', e);
                setStatus('Chyba pri otváraní zákazky.');
            } finally {
                opening = false;
            }
        }

        function matchesTerm(it, term) {
            const dashIndex = term.indexOf('-');
            if (dashIndex !== -1) {
                const orderPart = term.slice(0, dashIndex).trim();
                const opPart = term.slice(dashIndex + 1).trim();
                const orderMatch = !orderPart || (it.productionOrderNo || '').toLowerCase().includes(orderPart);
                const opMatch = !opPart || (it.operationNo || '').toLowerCase().includes(opPart);
                return orderMatch && opMatch;
            }
            return (
                (it.productionOrderNo || '').toLowerCase().includes(term) ||
                (it.salesOrderNo || '').toLowerCase().includes(term) ||
                (it.materialNo || '').toLowerCase().includes(term) ||
                (it.material || '').toLowerCase().includes(term)
            );
        }

        function buildSearchUI(sidebar) {
            const list = document.createElement('div');
            list.id = UI_ID;
            list.className = 'sapMList sapMListBGSolid';
            list.style.width = '100%';

            const header = document.createElement('div');
            header.className = 'sapMIBar sapMTB sapMTBNewFlex sapMTBInactive sapMTBStandard sapMTB-Transparent-CTX sapMListHdr sapMListHdrTBar sapMTBHeader-CTX';
            const headerTitle = document.createElement('div');
            headerTitle.className = 'sapMTitle sapMTitleStyleAuto sapMTitleNoWrap sapUiSelectable sapMTitleMaxWidth sapMTitleTB sapMBarChild sapMTBShrinkItem';
            headerTitle.innerHTML = '<span dir="auto">Číslo zákazky</span>';
            header.appendChild(headerTitle);

            const tbContainer = document.createElement('div');
            tbContainer.className = 'sapMListInfoTBarContainer';
            const toolbar = document.createElement('div');
            toolbar.className = 'sapMIBar sapMTB sapMTBNewFlex sapMTBInactive sapMTBClear sapMTB-Transparent-CTX sapMListInfoTBar';
            const sf = document.createElement('div');
            sf.className = 'sapMSF sapMSFVal sapMBarChild sapMTBShrinkItem';
            sf.style.width = '100%';

            const form = document.createElement('form');
            form.className = 'sapMSFF';
            form.addEventListener('submit', (e) => e.preventDefault());

            const input = document.createElement('input');
            input.type = 'search';
            input.autocomplete = 'off';
            input.placeholder = 'Číslo zákazky / materiál';
            input.className = 'sapMSFI';

            const resetDiv = document.createElement('div');
            resetDiv.title = 'Resetovať';
            resetDiv.className = 'sapMSFR sapMSFB';
            resetDiv.addEventListener('click', () => { input.value = ''; render(''); });

            const searchDiv = document.createElement('div');
            searchDiv.title = 'Hľadať';
            searchDiv.className = 'sapMSFS sapMSFB';

            form.appendChild(input);
            form.appendChild(resetDiv);
            form.appendChild(searchDiv);
            sf.appendChild(form);
            toolbar.appendChild(sf);
            tbContainer.appendChild(toolbar);

            const status = document.createElement('div');
            status.id = '__pda_status__';
            status.style.fontSize = '0.72rem';
            status.style.color = '#888';
            status.style.padding = '0.2rem 0.5rem';

            const resultsUl = document.createElement('ul');
            resultsUl.className = 'sapMListItems sapMListUl sapMListHighlight sapMListShowSeparatorsAll sapMListModeSingleSelectMaster';
            resultsUl.setAttribute('role', 'listbox');
            resultsUl.tabIndex = 0;

            function renderItem(it) {
                const li = document.createElement('li');
                li.tabIndex = 0;
                li.setAttribute('role', 'option');
                li.className = 'sapMLIB sapMLIB-CTX sapMLIBShowSeparator sapMLIBTypeActive sapMLIBActionable sapMLIBHoverable sapMLIBFocusable sapMCLI sapUiTinyMargin';
                li.innerHTML =
                    '<div class="sapMLIBContent">' +
                    '<div class="sapMFlexBoxFit sapMFlexBox sapMHBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent" style="height:100%;width:100%;">' +
                    '<div class="sapMFlexBox sapMVBox sapMFlexBoxJustifyStart sapMFlexBoxAlignItemsStretch sapMFlexBoxWrapNoWrap sapMFlexBoxAlignContentSpaceBetween sapMFlexBoxBGTransparent sapMFlexItem" style="height:100%;width:100%;">' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="font-weight:bold;text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + formatProductionOrder(it) + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.workcenter || '') + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapUiTinyMargin sapUiNoMarginBottom sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.material || '') + '</bdi></span></span>' +
                    '<span class="sapMLabel sapUiSelectable sapMLabelMaxWidth sapMFlexItem" style="text-align:left;"><span class="sapMLabelTextWrapper"><bdi>' + (it.descriptionShort || '') + '</bdi></span></span>' +
                    '</div></div></div>';

                li.addEventListener('click', () => openItem(it));
                li.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openItem(it); }
                });
                return li;
            }

            function idleStatus(index) {
                return 'Index: ' + index.length + ' zákaziek' +
                    (shared.ordersIndexUpdatedAt
                        ? ' (aktualiz. ' + shared.ordersIndexUpdatedAt.toLocaleTimeString() + ')'
                        : ' (čaká sa na načítanie aplikácie)');
            }

            function render(filterText) {
                resultsUl.innerHTML = '';
                const term = (filterText || '').trim().toLowerCase();
                const index = shared.ordersIndex || [];

                if (!term) {
                    setStatus(idleStatus(index));
                    lastAutoOpenedKey = null;
                    return;
                }

                const matches = index.filter((it) => matchesTerm(it, term));
                setStatus(matches.length + ' výsledok/-ov (z ' + index.length + ' položiek)');
                matches.slice(0, 200).forEach((it) => resultsUl.appendChild(renderItem(it)));

                if (matches.length === 1) {
                    const key = getItemKey(matches[0]);
                    if (key !== lastAutoOpenedKey) {
                        lastAutoOpenedKey = key;
                        openItem(matches[0]);
                        setTimeout(() => {
                            input.value = '';
                            resultsUl.innerHTML = '';
                            setStatus(idleStatus(index));
                            lastAutoOpenedKey = null;
                        }, 500);
                    }
                } else {
                    lastAutoOpenedKey = null;
                }
            }

            input.addEventListener('input', () => render(input.value));
            input.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                const decoded = padOperationPart(decodeScannerInput(input.value));
                input.value = decoded;
                render(decoded);
            });

            list.appendChild(header);
            list.appendChild(tbContainer);
            list.appendChild(status);
            list.appendChild(resultsUl);
            sidebar.appendChild(list);

            render('');
            renderFn = render;

            // sprístupní vyplnenie vyhľadávania pre modul skenera
            shared.fillSearchInput = (value) => {
                input.value = value;
                render(value);
            };
        }

        function ensureLayout() {
            if (opening) return;
            if (!document.getElementById(HOME_BUTTON_ID) || !document.getElementById(CONTENT_ID)) return;

            const parentSection = document.getElementById(PARENT_SECTION_ID);
            const tasksPanel = document.getElementById(TASKS_PANEL_ID);
            const workcenterPanel = document.getElementById(PANEL_ID);
            if (!parentSection || !tasksPanel || !workcenterPanel) return;

            let container = document.getElementById(SIDEBAR_ID);
            if (!container) {
                container = document.createElement('div');
                container.id = SIDEBAR_ID;
                container.style.width = '100%';
                container.style.boxSizing = 'border-box';
                container.style.maxHeight = '80vh';
                container.style.overflowY = 'auto';
                container.style.margin = '10px 0';
                container.style.backgroundColor = '#ffffff';
                container.style.borderRadius = '10px';
            }

            if (container.nextElementSibling !== workcenterPanel || container.parentElement !== parentSection) {
                parentSection.insertBefore(container, workcenterPanel);
            }

            if (!document.getElementById(UI_ID)) buildSearchUI(container);
        }

        DomWatch.add(ensureLayout);
    }

    /* ------------------ 3.5 Bocny panel pouzivatelov -------------------- */

    function modUsersPanel() {
        const PARENT_ID = 'Main';
        const CONTENT_ID = 'Main--MainPage';
        const WRAPPER_ID = '__pda_main_wrapper__';
        const PANEL_ID = '__pda_user_switch_panel__';
        const STYLE_TAG_ID = '__pda_panel_styles__';
        const PANEL_WIDTH = '280px';

        function injectStyles() {
            if (document.getElementById(STYLE_TAG_ID)) return;
            const style = document.createElement('style');
            style.id = STYLE_TAG_ID;
            style.textContent =
                '.pda-user-btn { border: 2px solid transparent; }' +
                '.pda-user-btn:hover { border: 2px solid #e9e9f0; }';
            document.head.appendChild(style);
        }

        function buildUserButtons(panel) {
            if (settings.users.length === 0) {
                const hint = document.createElement('div');
                hint.textContent = 'Zatiaľ nie sú zadaní žiadni používatelia. Doplň ich v nastaveniach (ozubené koliesko vpravo dole).';
                hint.style.color = '#c9c9dd';
                hint.style.fontSize = '0.8rem';
                hint.style.lineHeight = '1.4';
                panel.appendChild(hint);
                return;
            }

            settings.users.forEach(({ username, password }) => {
                const btn = document.createElement('button');
                btn.textContent = username;
                btn.classList.add('pda-user-btn');
                btn.style.display = 'block';
                btn.style.width = '100%';
                btn.style.padding = '0.5rem 0.8rem';
                btn.style.marginBottom = '0.4rem';
                btn.style.cursor = 'pointer';
                btn.style.textAlign = 'left';
                btn.style.borderRadius = '8px';
                btn.style.color = '#e9e9f0';
                btn.style.backgroundColor = '#222252';
                btn.addEventListener('click', () => UserSwitch.to(username, password));
                panel.appendChild(btn);
            });
        }

        function ensureLayout() {
            const parent = document.getElementById(PARENT_ID);
            const content = document.getElementById(CONTENT_ID);
            if (!parent || !content) return;

            let wrapper = document.getElementById(WRAPPER_ID);
            let panel = document.getElementById(PANEL_ID);
            if (wrapper && wrapper.contains(content) && panel) return;

            if (!wrapper) {
                wrapper = document.createElement('div');
                wrapper.id = WRAPPER_ID;
                wrapper.style.display = 'flex';
                wrapper.style.flexDirection = 'row';
                wrapper.style.alignItems = 'stretch';
                wrapper.style.width = '100%';
                wrapper.style.height = '100%';
            }

            if (!panel) {
                panel = document.createElement('div');
                panel.id = PANEL_ID;
                panel.style.flex = '0 0 ' + PANEL_WIDTH;
                panel.style.maxWidth = PANEL_WIDTH;
                panel.style.marginRight = '4px';
                panel.style.background = '#313175';
                panel.style.overflowY = 'auto';
                panel.style.padding = '0.75rem';
                panel.style.boxSizing = 'border-box';

                const title = document.createElement('div');
                title.textContent = 'Rýchla zmena používateľa';
                title.style.fontWeight = 'bold';
                title.style.marginBottom = '0.5rem';
                title.style.color = '#e9e9f0';
                panel.appendChild(title);

                buildUserButtons(panel);
            }

            if (!wrapper.contains(content)) {
                parent.insertBefore(wrapper, content);
                wrapper.appendChild(panel);
                wrapper.appendChild(content);
                content.style.flex = '1 1 auto';
                content.style.minWidth = '0';
            }
        }

        onReady(injectStyles);
        DomWatch.add(ensureLayout);
    }

    /* ------------------- 3.6 Skener a RFID citacka ---------------------- */

    function modInputListener() {
        const HIDDEN_INPUT_ID = '__pda_hidden_scanner_input__';
        const CARD_ID_LENGTH = 10;
        const REFOCUS_INTERVAL = 1000;

        function handleRawInput(raw) {
            const decoded = decodeScannerInput(raw);

            if (decoded.length === CARD_ID_LENGTH) {
                const user = UserSwitch.findByCardId(decoded);
                if (user) {
                    console.log(LOG, 'karta patrí používateľovi:', user.username);
                    UserSwitch.to(user.username, user.password);
                } else {
                    console.warn(LOG, 'karta nerozpoznaná, ID:', decoded);
                }
                return;
            }

            const finalValue = padOperationPart(decoded);
            console.log(LOG, 'rozpoznané číslo zákazky:', finalValue);
            if (typeof shared.fillSearchInput === 'function') {
                shared.fillSearchInput(finalValue);
            } else {
                console.warn(LOG, 'modul vyhľadávania nebeží, zákazka nebola vložená');
            }
        }

        function createHiddenInput() {
            let input = document.getElementById(HIDDEN_INPUT_ID);
            if (input) return input;

            input = document.createElement('input');
            input.id = HIDDEN_INPUT_ID;
            input.type = 'text';
            input.autocomplete = 'off';
            input.style.position = 'fixed';
            input.style.top = '0';
            input.style.left = '0';
            input.style.width = '1px';
            input.style.height = '1px';
            input.style.opacity = '0';
            input.style.border = 'none';
            input.style.padding = '0';
            input.style.margin = '0';
            input.style.pointerEvents = 'none';

            input.addEventListener('keydown', (e) => {
                if (e.key !== 'Enter') return;
                e.preventDefault();
                const raw = input.value;
                input.value = '';
                if (raw) handleRawInput(raw);
            });

            document.body.appendChild(input);
            return input;
        }

        function ensureFocus(input) {
            const active = document.activeElement;
            const isButton = !!active && (active.tagName === 'BUTTON' || active.getAttribute('role') === 'button');
            // nekradne focus, ked pouzivatel pise do policka alebo je otvoreny panel nastaveni
            if (document.getElementById('__pda_settings_overlay__') || document.getElementById('__pda_settings_pass__')) return;
            if (!active || active === document.body || isButton) input.focus();
        }

        onReady(() => {
            const input = createHiddenInput();
            input.focus();
            setInterval(() => ensureFocus(input), REFOCUS_INTERVAL);
        });
    }

    /* ---------------------- 3.7 Tlacidlo vykresu ------------------------ */

    function modDrawingButton() {
        const DB_NAME = 'pda_drawing_db';
        const STORE_NAME = 'handles';
        const HANDLE_KEY = 'excel_file_handle';

        // stlpce sa daju prestavit v nastaveniach (predvolene H, AH, AI)
        const COL_ORDER_NO = colToIndex(settings.excel.colOrder, 7);
        const COL_DRAWING_NO = colToIndex(settings.excel.colDrawing, 33);
        const COL_VERSION = colToIndex(settings.excel.colVersion, 34);

        const CONTAINER_ID = 'WorkcenterDetail--Order_FlexBox';
        const WRAPPER_ID = '__pda_order_drawing_wrapper__';
        const BUTTON_ID = '__pda_order_drawing_button__';
        const LOAD_BUTTON_ID = '__pda_order_drawing_load_button__';

        let drawingIndex = {};
        let currentDrawingInfo = null;
        let pendingHandle = null;
        let currentLoadState = 'checking';

        const supportsFsAccess = typeof W.showOpenFilePicker === 'function';

        // --- IndexedDB (aby vyber suboru prezil obnovenie stranky) ---

        function openHandleDb() {
            return new Promise((resolve, reject) => {
                const req = W.indexedDB.open(DB_NAME, 1);
                req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
                req.onsuccess = () => resolve(req.result);
                req.onerror = () => reject(req.error);
            });
        }

        async function saveHandle(handle) {
            const db = await openHandleDb();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, 'readwrite');
                tx.objectStore(STORE_NAME).put(handle, HANDLE_KEY);
                tx.oncomplete = () => resolve();
                tx.onerror = () => reject(tx.error);
                tx.onabort = () => reject(tx.error);
            });
        }

        async function loadHandle() {
            const db = await openHandleDb();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, 'readonly');
                const req = tx.objectStore(STORE_NAME).get(HANDLE_KEY);
                req.onsuccess = () => resolve(req.result || null);
                req.onerror = () => reject(req.error);
            });
        }

        // --- Excel ---

        /*
         * Cislo vyrobnej zakazky sa v Exceli a v PDA nepise rovnako:
         *   PDA:   001600089585
         *   Excel: '1600089585   (apostrof = textova bunka, bez uvodnych nul)
         * Povodny skript to riesil tak, ze natvrdo odrezal prve dva znaky a
         * pridal apostrof - staci mala zmena formatu v Exceli a nenajde nic.
         * Preto sa obe strany prevedu na rovnaky tvar: bez apostrofu, bez
         * medzier a bez uvodnych nul.
         */
        // '001600108008' aj "'1600108008" -> '1600108008' (len cislice, bez uvodnych nul)
        function normalizeOrderKey(value) {
            return String(value === undefined || value === null ? '' : value)
                .replace(/\D/g, '')
                .replace(/^0+/, '');
        }

        // Stlpce sa hladaju podla NAZVU v hlavicke (AutomatedOQ180.xlsx ich ma
        // stale rovnake), pismena z nastaveni su len zaloha pre iny subor.
        const HEADER_NAMES = {
            order: 'production order number',
            drawing: 'document number (material of production order)',
            version: 'document version (material of production order)',
            material: 'material number (production order)',
        };

        // 'SIEHE DIS', 'B.V.', 'O.Z.' nie su cisla vykresov - cislo vykresu ma vzdy cislicu
        function isDrawingNo(v) {
            return /\d/.test(String(v || ''));
        }

        function cellText(row, col) {
            if (col < 0 || row[col] === undefined || row[col] === null) return '';
            return String(row[col]).trim().replace(/^'+/, '');
        }

        function findColumns(headerRow) {
            const hdr = (headerRow || []).map((h) => String(h === undefined || h === null ? '' : h).trim().toLowerCase());
            const cols = {
                order: hdr.indexOf(HEADER_NAMES.order),
                drawing: hdr.indexOf(HEADER_NAMES.drawing),
                version: hdr.indexOf(HEADER_NAMES.version),
                material: hdr.indexOf(HEADER_NAMES.material),
            };
            if (cols.order === -1 || cols.drawing === -1) {
                console.log(LOG, 'Excel: hlavičky stĺpcov sa nenašli, používam písmená z nastavení');
                return { order: COL_ORDER_NO, drawing: COL_DRAWING_NO, version: COL_VERSION, material: -1 };
            }
            return cols;
        }

        function buildDrawingIndex(rows) {
            const cols = findColumns(rows[0]);
            const index = {};
            for (let i = 1; i < rows.length; i++) {
                const row = rows[i];
                if (!row) continue;
                const key = normalizeOrderKey(row[cols.order]);
                if (!key || index[key]) continue;   // prvy vyskyt vyhrava (v Python verzii overene: 0 konfliktov)
                const vyk = cellText(row, cols.drawing);
                const valid = isDrawingNo(vyk);
                index[key] = {
                    drawingNo: valid ? vyk : '',
                    version: valid ? cellText(row, cols.version) : '',
                    materialNo: cellText(row, cols.material).replace(/^0+/, ''),
                };
            }
            return index;
        }

        function parseWorkbook(bytes) {
            const workbook = XLSX.read(bytes, { type: 'array' });
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            if (!sheet) {
                console.warn(LOG, 'v exceli sa nenašiel žiadny list');
                return false;
            }
            drawingIndex = buildDrawingIndex(XLSX.utils.sheet_to_json(sheet, { header: 1 }));
            console.log(LOG, 'index výkresov vytvorený, záznamov:', Object.keys(drawingIndex).length);
            updateLoadButtonState('loaded');
            if (shared.currentOperation) applyForOperation(shared.currentOperation);
            return true;
        }

        async function loadExcelFromFile(file) {
            parseWorkbook(new Uint8Array(await file.arrayBuffer()));
        }

        /*
         * Zdroj Excelu z nastaveni:
         *   http(s)://...          -> stiahne sa zo servera
         *   C:\...  alebo  \\server\... alebo file:///... -> precita sa z disku cez file://
         *      (rovnako ako v Python appke; Tampermonkey na to potrebuje zapnute
         *       "Povolit pristup k URL adresam suborov" na stranke chrome://extensions)
         *   prazdne                -> rucny vyber suboru, prehliadac si ho zapamata
         */
        function excelSource() {
            let raw = String(settings.excel.url || '').trim().replace(/^"+|"+$/g, '');
            if (!raw) return { kind: 'none', url: '', raw: '' };
            if (/^https?:\/\//i.test(raw)) return { kind: 'http', url: raw, raw };
            if (/^file:\/\//i.test(raw)) return { kind: 'file', url: raw, raw };
            if (/^\\\\/.test(raw)) {                       // \\server\share\subor -> file://server/share/subor
                return { kind: 'file', url: 'file:' + encodeURI(raw.replace(/\\/g, '/')), raw };
            }
            if (/^[a-zA-Z]:[\\/]/.test(raw)) {               // C:\priecinok\subor -> file:///C:/priecinok/subor
                return { kind: 'file', url: 'file:///' + encodeURI(raw.replace(/\\/g, '/')), raw };
            }
            return { kind: 'invalid', url: '', raw };
        }

        // Excel stiahnuty z adresy (http://... alebo file://...). GM_xmlhttpRequest obchadza CORS,
        // takze sa da siahnut aj na interny server; pri file:// vracia Chrome status 0.
        function fetchExcel(url) {
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url,
                    responseType: 'arraybuffer',
                    onload: (r) => {
                        const hasData = r.response && r.response.byteLength > 0;
                        const ok = (r.status >= 200 && r.status < 300 && hasData) || (r.status === 0 && hasData);
                        if (!ok) {
                            reject(new Error(r.status === 0
                                ? 'prehliadač súbor nevydal — chýba povolenie prístupu k súborom?'
                                : 'HTTP ' + r.status));
                            return;
                        }
                        resolve(r.response);
                    },
                    onerror: () => reject(new Error(/^file:/i.test(url)
                        ? 'súbor sa nedá otvoriť — cesta alebo povolenie prístupu k súborom'
                        : 'spojenie so serverom zlyhalo')),
                    ontimeout: () => reject(new Error('zdroj neodpovedal včas')),
                });
            });
        }

        async function loadExcelFromUrl(url) {
            updateLoadButtonState('loading');
            try {
                const buffer = await fetchExcel(url);
                if (!buffer) throw new Error('prázdna odpoveď');
                parseWorkbook(new Uint8Array(buffer));
            } catch (e) {
                console.warn(LOG, 'Excel sa nepodarilo načítať z', url, e);
                updateLoadButtonState(/^file:/i.test(url) ? 'file-denied' : 'url-error');
            }
        }

        async function pickFileAndRemember() {
            try {
                const [handle] = await W.showOpenFilePicker({
                    types: [{
                        description: 'Excel',
                        accept: {
                            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                            'application/vnd.ms-excel': ['.xls'],
                        },
                    }],
                    multiple: false,
                });
                await saveHandle(handle);
                updateLoadButtonState('loading');
                await loadExcelFromFile(await handle.getFile());
            } catch (err) {
                if (err && err.name === 'AbortError') return;
                console.warn(LOG, 'chyba pri výbere súboru', err);
                updateLoadButtonState('error');
            }
        }

        async function tryAutoLoad() {
            if (!supportsFsAccess) { updateLoadButtonState('unsupported'); return; }

            let handle = null;
            try { handle = await loadHandle(); } catch (e) { /* ignore */ }
            if (!handle) { updateLoadButtonState('nofile'); return; }

            let permission;
            try {
                permission = await handle.queryPermission({ mode: 'read' });
            } catch (e) {
                updateLoadButtonState('needs-permission', handle);
                return;
            }

            if (permission === 'granted') {
                updateLoadButtonState('loading');
                try { await loadExcelFromFile(await handle.getFile()); }
                catch (e) { updateLoadButtonState('error'); }
                return;
            }
            updateLoadButtonState('needs-permission', handle);
        }

        async function confirmPermissionAndLoad(handle) {
            try {
                const permission = await handle.requestPermission({ mode: 'read' });
                if (permission !== 'granted') { updateLoadButtonState('needs-permission', handle); return; }
                updateLoadButtonState('loading');
                await loadExcelFromFile(await handle.getFile());
            } catch (e) {
                updateLoadButtonState('error');
            }
        }

        // --- napojenie na aktualne otvorenu operaciu ---

        /*
         * Prave otvorena operacia sa cita priamo z premennej appky
         * getGlobals().getVar('oSelectedWorkcenterOperation') - rovnako ako v
         * Python verzii. Povodne odpocuvanie siete cakalo na result.operation,
         * ktore appka neposiela spolahlivo, preto tlacidlo ostavalo prazdne.
         */
        function readSelectedOperation() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                const op = main && main.getController().getGlobals().getVar('oSelectedWorkcenterOperation');
                if (!op || !op.productionOrderNo) return null;
                return {
                    workcenter: op.workcenterDescription,
                    workcenterCode: op.workcenter,
                    productionOrderNo: String(op.productionOrderNo || ''),
                    salesOrderNo: String(op.salesOrderNo || ''),
                    operationNo: String(op.operationNo || ''),
                    sequenceNo: op.sequenceNo,
                    materialNo: String(op.materialNo || '').replace(/^0+/, ''),
                    material: op.material,
                };
            } catch (e) {
                return null;
            }
        }

        /*
         * Mapa vyrobna zakazka -> { kluc (cislo vykresu alebo material), typ, revizia }.
         * Stiahne sa zo sluzby RAZ pri starte (endpoint /zakazky = unikatne zakazky z Excelu,
         * prvy vyskyt vyhrava) a drzi sa v pamati aj v ulozisku Tampermonkey, aby prezila
         * obnovenie stranky. Znovu sa stahuje len ked sa na serveri zmeni odtlacok mapy.
         * Do Excelu sa z prehliadaca nechodi vobec - cita ho server raz denne.
         */
        const KEY_ORDER_MAP = 'pda_zakazky_cache_v1';
        let orderMap = null;
        let orderMapStamp = '';
        let orderMapLoading = false;

        (function restoreOrderMap() {
            const c = loadJson(KEY_ORDER_MAP, null);
            if (c && c.zakazky) {
                orderMap = c.zakazky;
                orderMapStamp = String(c.mapa_z || '');
                console.log(LOG, 'mapa zákaziek z úložiska:', Object.keys(orderMap).length, 'zákaziek (mapa z ' + orderMapStamp + ')');
            }
        })();

        function pdmGetJson(path, timeout) {
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + path,
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    timeout: timeout || 15000,
                    onload: (r) => {
                        if (r.status === 404) { resolve(null); return; }
                        if (r.status < 200 || r.status >= 300) { reject(new Error('HTTP ' + r.status + ' ' + path)); return; }
                        try { resolve(JSON.parse(r.responseText)); }
                        catch (e) { reject(new Error('neplatná odpoveď ' + path)); }
                    },
                    onerror: () => reject(new Error('spojenie so službou zlyhalo')),
                    ontimeout: () => reject(new Error('služba neodpovedala včas')),
                });
            });
        }

        async function loadOrderMap() {
            if (orderMapLoading) return;
            orderMapLoading = true;
            try {
                // lacna kontrola odtlacku: /health ma ~400 B, cela mapa stovky kB
                const h = await pdmGetJson('/health', 8000);
                const stamp = h && h.mapa && h.mapa.posledne ? String(h.mapa.posledne) : '';
                if (orderMap && stamp && stamp === orderMapStamp) {
                    console.log(LOG, 'mapa zákaziek je aktuálna (' + stamp + ')');
                    return;
                }
                const d = await pdmGetJson('/zakazky', 30000);
                if (!d || !d.zakazky) {
                    console.log(LOG, 'služba ešte nemá /zakazky — hľadám po jednej zákazke');
                    return;
                }
                orderMap = d.zakazky;
                orderMapStamp = String(d.mapa_z || stamp || '');
                saveJson(KEY_ORDER_MAP, { mapa_z: orderMapStamp, zakazky: orderMap });
                console.log(LOG, 'mapa zákaziek stiahnutá:', d.pocet, 'zákaziek (mapa z ' + orderMapStamp + ')');
                if (shared.currentOperation) applyForOperation(shared.currentOperation);
            } catch (e) {
                console.log(LOG, 'mapa zákaziek sa nestiahla:', e.message);
            } finally {
                orderMapLoading = false;
            }
        }

        let lastOperationKey = null;

        function pollSelectedOperation() {
            const cur = readSelectedOperation();
            const key = cur ? cur.productionOrderNo + '|' + cur.operationNo + '|' + cur.sequenceNo : '';
            if (key === lastOperationKey) return;
            lastOperationKey = key;
            shared.currentOperation = cur;
            if (cur) applyForOperation(cur);
            else updateDisplay(null);
        }

        onReady(loadOrderMap);
        onReady(() => setInterval(pollSelectedOperation, 500));

        function stripLeadingApostrophe(value) {
            if (!value) return value;
            return value.charAt(0) === "'" ? value.slice(1) : value;
        }

        function salesOrderOf(current) {
            return String((current && current.salesOrderNo) || '').trim();
        }

        // "3-55.2-06.74-075_001_A_0.pdf" -> "3-55.2-06.74-075"
        function drawingNoFromResult(v) {
            if (v.cislo_vykresu) return v.cislo_vykresu;
            if (v.kluc && v.kluc_typ === 'vykres') return v.kluc;
            return String(v.nazov || '').replace(/\.(pdf|tiff?)$/i, '').split('_')[0];
        }

        // Ked sluzba vrati viac suborov, najrelevantnejsi je ten, ktory lezi
        // v priecinku tejto zakazky a ktoremu sedi revizia so SAP verziou.
        function bestResult(list) {
            return list.find((v) => v.v_zakazke && v.zhoda_revizie === true) ||
                   list.find((v) => v.v_zakazke) ||
                   list.find((v) => v.zhoda_revizie === true) ||
                   list[0] || null;
        }

        let lookupToken = 0;

        /*
         * Cislo vykresu sa zistuje v dvoch krokoch:
         *   1. z Excelu podla cisla vyrobnej zakazky (ak je Excel nacitany)
         *   2. inak priamo zo sluzby vykresov podla cisla materialu, ktore
         *      PDA uz pozna - Excel teda nie je podmienkou
         * Zakaznicka zakazka ide do parametra "path", vdaka comu sluzba oznaci
         * vykresy lezice priamo v tejto zakazke (v_zakazke = true).
         */
        async function applyForOperation(current) {
            const token = ++lookupToken;

            const fromExcel = drawingIndex[normalizeOrderKey(current.productionOrderNo)];
            const excelDrawingNo = fromExcel ? stripLeadingApostrophe(fromExcel.drawingNo) : '';
            if (excelDrawingNo) {
                updateDisplay({
                    drawingNo: excelDrawingNo,
                    version: stripLeadingApostrophe(fromExcel.version) || '',
                    searchTerm: excelDrawingNo,
                    source: 'excel',
                });
                return;
            }

            updateDisplay({ drawingNo: '…', version: '', searchTerm: '', source: 'hladam' });

            // 2) sluzba vykresov pozna denny export -> kluc priamo pre tuto vyrobnu zakazku
            try {
                const cislo = normalizeOrderKey(current.productionOrderNo);
                const local = orderMap && cislo ? orderMap[cislo] : null;
                const z = local
                    ? { kluc: local.kluc, typ: local.typ, revizia: local.revizia, pocet: 0 }
                    : await pdmOrderLookup(current.productionOrderNo);
                if (token !== lookupToken) return;
                if (z && z.kluc) {
                    const isDrawing = z.typ !== 'material';
                    updateDisplay({
                        drawingNo: isDrawing ? z.kluc : 'SAP ' + z.kluc,
                        version: isDrawing ? (z.revizia || '') : '',
                        searchTerm: z.kluc,
                        source: 'server',
                        count: z.pocet || 0,
                    });
                    return;
                }
            } catch (e) {
                if (token !== lookupToken) return;
                console.log(LOG, 'služba nepozná zákazku (alebo ešte nemá /zakazka), skúšam materiál:', e.message);
            }

            // 3) zaloha: material - najprv z Excelu (ak zakazku pozna, ale vykres tam nema), inak z operacie
            const material = ((fromExcel && fromExcel.materialNo) || String(current.materialNo || '')).trim();
            if (!material) { updateDisplay(null); return; }

            updateDisplay({ drawingNo: '…', version: '', searchTerm: material, source: 'hladam' });

            try {
                const d = await pdmSearch(material, { path: salesOrderOf(current) });
                if (token !== lookupToken) return; // medzitym sa otvorila ina operacia

                const best = bestResult(d.vysledky || []);
                if (!best) { updateDisplay(null); return; }

                updateDisplay({
                    drawingNo: drawingNoFromResult(best),
                    version: best.revizia || '',
                    searchTerm: material,
                    source: 'pdm',
                    count: d.pocet || 0,
                });
            } catch (e) {
                if (token !== lookupToken) return;
                console.warn(LOG, 'hľadanie výkresu zlyhalo', e);
                updateDisplay(null);
            }
        }

        function updateDisplay(info) {
            const valueEl = document.getElementById('__pda_order_drawing_value__');
            const revisionEl = document.getElementById('__pda_order_drawing_revision__');
            if (!valueEl || !revisionEl) return;

            if (!info) {
                currentDrawingInfo = null;
                valueEl.textContent = '—';
                revisionEl.textContent = '';
                return;
            }

            currentDrawingInfo = info;
            valueEl.textContent = info.drawingNo || '—';

            if (info.source === 'hladam') {
                revisionEl.textContent = 'hľadám…';
            } else if (info.version) {
                revisionEl.textContent = 'rev. ' + info.version + (info.count > 1 ? ' · ' + info.count + ' súb.' : '');
            } else {
                revisionEl.textContent = info.count > 1 ? info.count + ' súbory' : '';
            }
        }

        // --- sluzba "Mapa vykresov" (PDM) ---

        /*
         * Vyrobna zakazka -> kluc z denneho exportu. Sluzba vykresov cita AutomatedOQ180.xlsx
         * raz denne a ku kazdemu vykresu drzi zoznam zakaziek; endpoint /zakazka/<cislo>
         * to vrati jednym dopytom. 404 = zakazka v exporte nie je.
         */
        function pdmOrderLookup(productionOrderNo) {
            const cislo = normalizeOrderKey(productionOrderNo);
            if (!cislo) return Promise.resolve(null);
            return new Promise((resolve, reject) => {
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + '/zakazka/' + encodeURIComponent(cislo),
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    timeout: 15000,
                    onload: (r) => {
                        if (r.status === 404) { resolve(null); return; }
                        if (r.status < 200 || r.status >= 300) { reject(new Error('HTTP ' + r.status)); return; }
                        try { resolve(JSON.parse(r.responseText)); }
                        catch (e) { reject(new Error('neplatná odpoveď služby')); }
                    },
                    onerror: () => reject(new Error('spojenie so službou zlyhalo')),
                    ontimeout: () => reject(new Error('služba neodpovedala včas')),
                });
            });
        }

        function pdmSearch(cislo, { path = '', live = false } = {}) {
            return new Promise((resolve, reject) => {
                const params = new URLSearchParams({ q: cislo });
                if (path) params.set('path', path);
                if (live) params.set('live', '1');
                GM_xmlhttpRequest({
                    method: 'GET',
                    url: settings.pdm.base + '/search?' + params,
                    headers: settings.pdm.key ? { 'X-API-Key': settings.pdm.key } : {},
                    onload: (response) => {
                        if (response.status < 200 || response.status >= 300) {
                            reject(new Error('Služba výkresov vrátila HTTP ' + response.status));
                            return;
                        }
                        try { resolve(JSON.parse(response.responseText)); }
                        catch (e) { reject(new Error('Neplatná odpoveď zo služby výkresov')); }
                    },
                    onerror: () => reject(new Error('Služba výkresov neodpovedala (chyba spojenia)')),
                    ontimeout: () => reject(new Error('Služba výkresov neodpovedala včas')),
                });
            });
        }

        /*
         * Zhoda revizie z Excelu s reviziou v nazve suboru - rovnake pravidlo ako
         * v sluzbe (Python verzia): pismeno musi sediet, index len ak ho poznaju
         * obaja. 'BC' ~ 'C' ano, 'BC' ~ 'B' nie, '0A' ~ 'A' ano.
         */
        function revisionMatches(label, rev) {
            label = String(label || '').toUpperCase();
            rev = String(rev || '').toUpperCase();
            if (!label || !rev || label === 'BEZ REV.') return false;
            const l1 = label.slice(-1), i1 = label.length > 1 ? label.slice(0, -1) : '';
            const l2 = rev.slice(-1), i2 = rev.length > 1 ? rev.slice(0, -1) : '';
            return l1 === l2 && (!i1 || !i2 || i1 === i2);
        }

        // Hlada sa LEN podla cisla vykresu (bez revizie); revizia z Excelu sa
        // v zozname iba zvyrazni, rozhodnutie ostava na cloveku.
        function pdmOpenDialog(cislo, { path = '', rev = '', titul = '' } = {}) {
            const overlay = document.createElement('div');
            overlay.id = '__pda_pdm_overlay__';
            overlay.style.cssText =
                'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99999;' +
                'display:flex;align-items:center;justify-content:center;font-family:inherit;';

            overlay.innerHTML =
                '<div style="background:#fff;border-radius:10px;max-width:900px;width:92%;max-height:80vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 8px 30px rgba(0,0,0,.3)">' +
                '<div style="display:flex;justify-content:space-between;align-items:center;padding:12px 16px;background:#5b9bd5;color:#fff">' +
                // v hlavicke je to, co vidi operator v okienku VYKRES (cislo vykresu
                // + revizia); hlada sa stale podla `cislo` (moze to byt aj material)
                '<b>Výkresy — ' + (titul || cislo) + '</b>' +
                '<button data-pda-zavrit type="button" style="background:none;border:0;color:#fff;font-size:22px;line-height:1;cursor:pointer">×</button>' +
                '</div><div data-pda-telo style="padding:14px 16px;overflow:auto">Hľadám…</div></div>';

            const close = () => { overlay.remove(); document.removeEventListener('keydown', onEsc); };
            const onEsc = (e) => { if (e.key === 'Escape') close(); };

            overlay.querySelector('[data-pda-zavrit]').addEventListener('click', close);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
            document.addEventListener('keydown', onEsc);
            document.body.appendChild(overlay);

            const telo = overlay.querySelector('[data-pda-telo]');

            const TH = 'style="text-align:left;padding:6px;border-bottom:2px solid #5b9bd5"';
            const TD = 'style="padding:6px;border-bottom:1px solid #eee"';

            function run(live) {
                telo.innerHTML = live
                    ? '<p>Hľadám naživo v PDM… (môže to trvať pár sekúnd)</p>'
                    : '<p>Hľadám…</p>';

                pdmSearch(cislo, { path, live }).then((d) => {
                    if (!d.pocet) {
                        telo.innerHTML =
                            '<p>Pre <b>' + cislo + '</b> sa nenašiel žiadny PDF ani TIFF výkres.</p>' +
                            (live ? '' :
                                '<p style="color:#777;font-size:12px">Denná mapa sa obnovuje raz za deň — čerstvo pridaný výkres v nej ešte nemusí byť.</p>' +
                                '<button data-pda-live type="button" style="background:#5b9bd5;color:#fff;border:0;border-radius:7px;padding:8px 15px;cursor:pointer;font:inherit;font-size:.88rem">Hľadať naživo v PDM</button>');
                        const liveBtn = telo.querySelector('[data-pda-live]');
                        if (liveBtn) liveBtn.addEventListener('click', () => run(true));
                        return;
                    }

                    // zhoda revizie: primarne zo sluzby (pocita ju zo SAP verzie), inak lokalne z Excelu
                    const rows = (d.vysledky || []).map((v) => {
                        const known = v.zhoda_revizie === true || v.zhoda_revizie === false;
                        return Object.assign({}, v, { zhoda: known ? v.zhoda_revizie : (rev ? revisionMatches(v.revizia, rev) : null) });
                    });
                    rows.sort((a, b) => ((b.zhoda === true) - (a.zhoda === true)) || String(a.nazov).localeCompare(String(b.nazov)));

                    telo.innerHTML =
                        '<table style="width:100%;border-collapse:collapse;font-size:14px"><thead><tr>' +
                        '<th ' + TH + '>Názov</th><th ' + TH + '>Revízia</th>' +
                        '<th ' + TH + '>Stav</th><th ' + TH + '>Ver.</th>' +
                        '</tr></thead><tbody>' +
                        rows.map((v) => {
                            const bg = v.zhoda === true ? '#e6f4ea' : (v.v_zakazke ? '#f3f7fd' : 'transparent');
                            const mark = v.v_zakazke ? ' <span style="font-size:10px;color:#2f6fbf;border:1px solid #b9cbe8;border-radius:9px;padding:1px 6px">v zákazke</span>' : '';
                            return '<tr data-id="' + v.id + '" style="cursor:pointer;background:' + bg + '">' +
                                '<td ' + TD + '><b>' + v.nazov + '</b>' + mark + '</td>' +
                                '<td ' + TD + '>' + (v.revizia || '') + '</td>' +
                                '<td ' + TD + '>' + (v.stav || '') + '</td>' +
                                '<td ' + TD + '>' + (v.verzia != null ? v.verzia : '') + '</td></tr>';
                        }).join('') +
                        '</tbody></table>' +
                        '<p data-pda-info style="color:#777;font-size:12px;margin-top:10px">' + d.pocet + ' výkresov · ' +
                        (d.zdroj === 'mapa' ? 'z dennej mapy' : 'naživo z PDM') +
                        (d.mapa_z ? ' (' + d.mapa_z + ')' : '') +
                        ' · zelený riadok = revízia sedí so SAP · kliknutím sa výkres otvorí</p>';

                    telo.querySelectorAll('tr[data-id]').forEach((tr) => {
                        tr.addEventListener('click', () => {
                            const info = telo.querySelector('[data-pda-info]');
                            // vydaj suboru trva 0,4-4 s, treba to dat najavo
                            if (info) info.textContent = 'Otváram výkres… (ťahá sa z PDM, môže to chvíľu trvať)';
                            W.open(settings.pdm.base + '/file/' + tr.dataset.id, '_blank');
                        });
                    });
                }).catch((err) => {
                    telo.innerHTML = '<p style="color:#b00">Služba výkresov neodpovedala.<br>' + err.message +
                        '<br><span style="color:#777;font-size:12px">Adresa: ' + settings.pdm.base + '</span></p>';
                });
            }

            run(false);
        }

        // --- tlacidla ---

        /*
         * Nadpis okna s vykresmi = to iste, co je v okienku VYKRES vpravo hore
         * ("1-60.2-06.66-009 rev. C"). Predtym tam bolo cislo, podla ktoreho sa
         * hlada - pri hladani podla materialu teda cislo materialu, co operatora
         * mylilo. Hladanie sa nemeni, meni sa len text v hlavicke.
         */
        function dialogTitle() {
            if (!currentDrawingInfo) return '';
            const cislo = String(currentDrawingInfo.drawingNo || '').trim();
            if (!cislo || cislo === '—' || cislo === '…') return '';
            const rev = String(currentDrawingInfo.version || '').trim();
            return rev ? cislo + ' rev. ' + rev : cislo;
        }

        /*
         * Hlasenia o Exceli sa uz operatorovi neukazuju. Vykres sa najde aj bez
         * Excelu (cez sluzbu PDM podla materialu), takze cervene "Excel sa
         * nestiahol - skus znova" pri otvorenej zakazke len zavadzalo. Jedina
         * vynimka je stav "needs-permission": tam prehliadac ziada kliknutie
         * cloveka, bez tlacidla by sa Excel nedal nacitat vobec.
         * Ostatne stavy idu uz len do konzoly.
         */
        function renderLoadButtonState(loadButton, state) {
            if (state !== 'needs-permission') {
                loadButton.style.display = 'none';
                if (state !== 'loaded' && state !== 'nofile' && state !== 'loading') {
                    console.log(LOG, 'stav Excelu:', state, '(hlasenie sa operatorovi nezobrazuje)');
                }
                return;
            }
            loadButton.style.display = '';

            const states = {
                loading: ['Načítavam…', '#f5f5f5', '#ccc'],
                'needs-permission': ['Povoliť prístup k Excelu', '#fff4e5', '#f9a825'],
                error: ['Chyba, skús znova', '#fdecea', '#e53935'],
                'url-error': ['Excel sa nestiahol — skús znova', '#fdecea', '#e53935'],
                'url-invalid': ['Adresa Excelu je nezrozumiteľná — klik otvorí nastavenia', '#fff4e5', '#f9a825'],
                'file-denied': ['Excel z disku sa nenačítal — zapni prístup k súborom v Tampermonkey', '#fdecea', '#e53935'],
                unsupported: ['Prehliadač nepodporuje zapamätanie', '#fdecea', '#e53935'],
                // Excel uz nie je podmienkou - vykres sa najde aj podla materialu
                nofile: ['Excel (nepovinné)', '#f5f5f5', '#ccc'],
            };
            const [text, bg, border] = states[state] || states.nofile;
            loadButton.textContent = text;
            loadButton.style.backgroundColor = bg;
            loadButton.style.borderColor = border;
        }

        function updateLoadButtonState(state, handle) {
            currentLoadState = state;
            pendingHandle = state === 'needs-permission' ? handle : null;
            const loadButton = document.getElementById(LOAD_BUTTON_ID);
            if (loadButton) renderLoadButtonState(loadButton, state);
        }

        function buildWrapper() {
            const wrapper = document.createElement('div');
            wrapper.id = WRAPPER_ID;
            wrapper.style.cssText = 'display:flex;justify-content:flex-end;align-items:center;width:100%;box-sizing:border-box;margin-bottom:8px;';

            const loadButton = document.createElement('button');
            loadButton.id = LOAD_BUTTON_ID;
            loadButton.type = 'button';
            loadButton.style.cssText = 'padding:6px 10px;border:1px solid #ccc;border-radius:6px;background:#f5f5f5;cursor:pointer;font-size:0.75rem;margin-right:8px;align-self:center;';
            loadButton.addEventListener('click', () => {
                const src = excelSource();
                if (src.kind === 'invalid') openSettings();
                else if (src.kind === 'http' || src.kind === 'file') loadExcelFromUrl(src.url);
                else if (pendingHandle) confirmPermissionAndLoad(pendingHandle);
                else pickFileAndRemember();
            });

            const button = document.createElement('button');
            button.id = BUTTON_ID;
            button.type = 'button';
            button.style.cssText = 'display:flex;flex-direction:column;align-items:flex-end;gap:2px;padding:6px 12px;border:1px solid #5b9bd5;border-radius:6px;background:#fff;cursor:pointer;width:150px;flex-shrink:0;margin-right:8px;';

            const label = document.createElement('span');
            label.textContent = 'VÝKRES';
            label.style.cssText = 'font-size:0.65rem;color:#888;letter-spacing:0.05em;';

            const value = document.createElement('span');
            value.id = '__pda_order_drawing_value__';
            value.textContent = '—';
            value.style.cssText = 'font-size:0.9rem;font-weight:bold;color:#000;';

            const revision = document.createElement('span');
            revision.id = '__pda_order_drawing_revision__';
            revision.style.cssText = 'font-size:0.75rem;color:#555;';

            button.appendChild(label);
            button.appendChild(value);
            button.appendChild(revision);
            button.addEventListener('click', () => {
                const path = salesOrderOf(shared.currentOperation);

                // ak uz vieme cislo vykresu (alebo aspon material), hladame podla neho
                if (currentDrawingInfo && currentDrawingInfo.searchTerm) {
                    pdmOpenDialog(currentDrawingInfo.searchTerm, {
                        path,
                        rev: (currentDrawingInfo.source === 'excel' || currentDrawingInfo.source === 'server')
                            ? currentDrawingInfo.version : '',
                        titul: dialogTitle(),
                    });
                    return;
                }
                // inak skusime aspon cislo materialu z otvorenej operacie
                const material = shared.currentOperation && String(shared.currentOperation.materialNo || '').trim();
                if (material) pdmOpenDialog(material, { path, titul: dialogTitle() });
                else console.log(LOG, 'pre aktuálnu zákazku nie je známe ani číslo výkresu, ani materiálu');
            });

            wrapper.appendChild(loadButton);
            wrapper.appendChild(button);
            return wrapper;
        }

        function ensureButton() {
            const container = document.getElementById(CONTAINER_ID);
            if (!container || document.getElementById(WRAPPER_ID)) return;

            container.insertBefore(buildWrapper(), container.firstChild);
            const loadButton = document.getElementById(LOAD_BUTTON_ID);
            if (loadButton) renderLoadButtonState(loadButton, currentLoadState);
            if (shared.currentOperation) applyForOperation(shared.currentOperation);
        }

        DomWatch.add(ensureButton);
        onReady(() => {
            // ak je v nastaveniach adresa, tahame odtial automaticky;
            // inak sa subor vybera rucne a prehliadac si ho pamata
            const src = excelSource();
            if (src.kind === 'http' || src.kind === 'file') {
                loadExcelFromUrl(src.url);
            } else if (src.kind === 'invalid') {
                console.warn(LOG, 'adresa Excelu je nezrozumiteľná, ignorujem:', src.raw);
                updateLoadButtonState('url-invalid');
            } else {
                tryAutoLoad();
            }
        });
    }

    /* ------------------ 3.8 Prehlad pracovisk (uvodna) ------------------ */

    /*
     * Portovane z Python appky (Temporary/PDA/app/pda_action.py, blok OBLASTI).
     * Rozdiel oproti prvej verzii tohto modulu:
     *  - data sa NEcitaju zo siete ani z DOM, ale priamo z premennej appky
     *    getGlobals().getVar('aWorkcenterOperationLists')
     *  - kategoria = pole `area` (Assembly / Welding / Machining / Quality Control)
     *  - "vyraba" = pole `count` > 0 (kolko ludi je prihlasenych na pracovisku)
     *  - klik na stroj vola ten isty handler, aky vola appka pri kliku na kartu
     *  - povodne karty sa skryju jednym CSS pravidlom, v DOM ostavaju
     */
    function modWorkcenterOverview() {
        const OVERVIEW_ID = '__pda_overview__';
        const STYLE_ID = '__pda_overview_styles__';
        const BODY_CLASS = 'pda-overview-on';
        const HOME_BUTTON_ID = 'Main--Button_HomeScreen';
        const PANEL_ID = 'Main--Workcenter_Panel';
        const ORDER = ['Assembly', 'Welding', 'Machining', 'Quality Control'];
        const SUBTITLES = {
            'Assembly': 'Finálna montáž a podzostavy',
            'Welding': 'Zváranie a príprava',
            'Machining': 'CNC a konvenčné obrábanie',
            'Quality Control': 'Kontrola a meranie',
            'Ostatné': 'Nezaradené pracoviská',
        };
        const ICONS = { 'Assembly': '🔧', 'Welding': '🔥', 'Machining': '⚙', 'Quality Control': '🔎' };

        const expanded = new Set();   // rozbalene kategorie
        const filters = {};           // { q, only } pre kazdu kategoriu
        let lastSignature = null;

        function globals() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                return main ? main.getController().getGlobals() : null;
            } catch (e) { return null; }
        }

        function workcenterList() {
            const g = globals();
            if (!g) return [];
            try { return g.getVar('aWorkcenterOperationLists') || []; } catch (e) { return []; }
        }

        function userName() {
            const g = globals();
            try {
                const n = g && g.getVar('oUser', 'username');
                if (n) return String(n);
            } catch (e) { /* ignore */ }
            const el = document.querySelector('[id$="Label_Username-bdi"]');
            return el ? el.textContent.trim() : '';
        }

        function normalize(s) {
            return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
        }

        /*
         * Otvorenie pracoviska presne tak, ako to robi appka po kliku na kartu
         * (handler f_cards_Header3_press). V Python verzii sa zistilo, ze handler
         * po navTo nedobehne - texty casov ostanu prazdne a potvrdzovacie tlacidla
         * zapnute - preto sa to po 1,5 s dorobi rovnako, ako to robi appka.
         */
        function openWorkcenter(wc) {
            try {
                const core = W.sap.ui.getCore();
                const main = core.byId('Main').getController();
                const g = main.getGlobals();
                const app = core.byId('Application');

                if (app && app.getCurrentPage && String(app.getCurrentPage().getId()).indexOf('WorkcenterDetail') === 0) {
                    main.navTo('Main', 'show');
                }
                if (!g.getVar('oSelectedWorkcenterOperation')) g.setVar('oSelectedWorkcenterOperation', {});

                const ctx = { getObject: () => wc };
                const src = { getParent: () => ({ getBindingContext: () => ctx }) };
                main.f_cards_Header3_press({ getSource: () => src });

                setTimeout(() => {
                    try {
                        const op = g.getVar('oSelectedWorkcenterOperation') || {};
                        [['SetupTime_Text', 'setupTimeConf', 'setupTime'],
                         ['MachineTime_Text', 'machineTimeConf', 'machineTime'],
                         ['LaborTime_Text', 'laborTimeConf', 'laborTime']].forEach(([id, conf, plan]) => {
                            const e = core.byId('WorkcenterDetail--' + id);
                            if (e && typeof e.setText === 'function') {
                                e.setText(op.id ? op[conf] + ' / ' + op[plan] + ' min' : '- / - min');
                            }
                        });
                        if (!op.id) {
                            main.setItemValue('WorkcenterDetail', 'Part_Confirm_Button', 'enabled', false);
                            main.setItemValue('WorkcenterDetail', 'Confirm_Button', 'enabled', false);
                        }
                    } catch (e) { /* kozmetika, nie je kriticka */ }
                }, 1500);
            } catch (e) {
                console.warn(LOG, 'otvorenie pracoviska cez handler zlyhalo, skúšam klik na kartu', e);
                const tile = findTile(wc);
                if (tile) pressElement(tile);
            }
        }

        // zaloha: povodna karta v DOM podla nazvu pracoviska
        function findTile(wc) {
            const want = String(wc.workcenterDescription || '').trim();
            const tiles = document.querySelectorAll('[id^="Main--Workcenter_Toolbar-Main--ui_layout_Grid3-"]');
            for (const t of tiles) {
                const suffix = t.id.replace('Main--Workcenter_Toolbar-', '');
                const el = document.getElementById('Main--WorkcenterDescription_Label-' + suffix + '-bdi');
                if (el && el.textContent.trim() === want) return t;
            }
            return null;
        }

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
body.${BODY_CLASS} #${PANEL_ID} .sapMPanelContent > :not(#${OVERVIEW_ID}) { display: none !important; }
#${OVERVIEW_ID} { font: 14px/1.5 -apple-system,"Segoe UI",Roboto,sans-serif; color:#1a2233; padding: 4px 0 10px; }
#${OVERVIEW_ID} .ov-hi { padding: 4px 6px 14px; }
#${OVERVIEW_ID} .ov-hi .k { font-size:.68rem; letter-spacing:.12em; color:#5b7cb5; font-weight:700; }
#${OVERVIEW_ID} .ov-hi .n { font-size:1.7rem; font-weight:700; margin:2px 0 1px; }
#${OVERVIEW_ID} .ov-hi .s { color:#6b7180; font-size:.87rem; }
#${OVERVIEW_ID} .ov-g { border:1px solid #dfe4ec; border-radius:12px; background:#fff; margin-bottom:10px; overflow:hidden; }
#${OVERVIEW_ID} .ov-g.open { border-color:#b9cbe8; box-shadow:0 1px 6px rgba(40,70,130,.08); }
#${OVERVIEW_ID} .ov-gh { display:flex; align-items:center; gap:14px; width:100%; background:none; border:0;
  padding:14px 16px; cursor:pointer; text-align:left; font:inherit; color:inherit; }
#${OVERVIEW_ID} .ov-g.open .ov-gh { background:#f2f6fc; }
#${OVERVIEW_ID} .ov-ic { width:42px; height:42px; border-radius:9px; background:#eaf1fb; flex-shrink:0;
  display:flex; align-items:center; justify-content:center; font-size:21px; }
#${OVERVIEW_ID} .ov-gt { flex:1 1 auto; min-width:0; }
#${OVERVIEW_ID} .ov-gt b { font-size:1.05rem; display:block; }
#${OVERVIEW_ID} .ov-gt span { color:#6b7180; font-size:.85rem; }
#${OVERVIEW_ID} .ov-num { text-align:center; flex-shrink:0; min-width:74px; }
#${OVERVIEW_ID} .ov-num b { display:block; font-size:1.3rem; color:#1d4ed8; line-height:1.1; }
#${OVERVIEW_ID} .ov-num span { font-size:.62rem; letter-spacing:.09em; color:#8b93a3; }
#${OVERVIEW_ID} .ov-num.on b { color:#2f7d43; }
#${OVERVIEW_ID} .ov-ch { color:#9aa3b4; font-size:18px; flex-shrink:0; }
#${OVERVIEW_ID} .ov-gb { padding:4px 16px 16px; border-top:1px solid #e8edf4; }
#${OVERVIEW_ID} .ov-tools { display:flex; gap:8px; justify-content:flex-end; margin:12px 0; flex-wrap:wrap; }
#${OVERVIEW_ID} .ov-tools input, #${OVERVIEW_ID} .ov-tools select {
  padding:6px 10px; border:1px solid #ccd4e0; border-radius:7px; font:inherit; font-size:.85rem; }
#${OVERVIEW_ID} .ov-tools input { width:190px; }
#${OVERVIEW_ID} .ov-chips { display:grid; grid-template-columns:repeat(auto-fill,minmax(215px,1fr)); gap:7px; }
#${OVERVIEW_ID} .ov-chip { display:flex; align-items:center; justify-content:space-between; gap:10px;
  border:1px solid #dfe4ec; border-radius:9px; background:#fff; padding:8px 11px; cursor:pointer;
  font:inherit; text-align:left; transition:border-color .12s, background .12s; }
#${OVERVIEW_ID} .ov-chip:hover { border-color:#5b8def; background:#f6f9ff; }
#${OVERVIEW_ID} .ov-chip b { font-size:.93rem; }
#${OVERVIEW_ID} .ov-txt { display:flex; flex-direction:column; min-width:0; }
#${OVERVIEW_ID} .ov-txt small { font-size:.71rem; color:#6b7180; line-height:1.3;
  white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
#${OVERVIEW_ID} .ov-st { display:flex; align-items:center; gap:5px; font-size:.76rem; color:#1e40af;
  background:#e8f0fe; border-radius:999px; padding:3px 9px; white-space:nowrap; font-weight:600; }
#${OVERVIEW_ID} .ov-dot { width:7px; height:7px; border-radius:50%; background:#2563eb; flex-shrink:0; }
#${OVERVIEW_ID} .ov-chip.run .ov-st { color:#166534; background:#eaf6ee; }
#${OVERVIEW_ID} .ov-chip.run .ov-dot { background:#16a34a; }
#${OVERVIEW_ID} .ov-empty { color:#8b93a3; font-size:.85rem; padding:8px 2px; }
`;
            document.head.appendChild(st);
        }

        function chip(w) {
            const producing = (w.count || 0) > 0;
            const code = String(w.workcenter || '');
            const name = String(w.workcenterDescription || '');

            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'ov-chip' + (producing ? ' run' : '');
            btn.title = code + ' — ' + name;

            const txt = document.createElement('span');
            txt.className = 'ov-txt';
            const label = document.createElement('b');
            label.textContent = code || name;
            txt.appendChild(label);
            if (code && name) {
                const sub = document.createElement('small');
                sub.textContent = name.length > 20 ? name.slice(0, 20).trim() + '…' : name;
                txt.appendChild(sub);
            }

            const st = document.createElement('span');
            st.className = 'ov-st';
            const dot = document.createElement('span');
            dot.className = 'ov-dot';
            st.appendChild(dot);
            st.appendChild(document.createTextNode(producing ? 'Vyrába' : 'Nevyrába'));

            btn.appendChild(txt);
            btn.appendChild(st);
            btn.addEventListener('click', () => openWorkcenter(w));
            return btn;
        }

        function groupBody(groupName, items) {
            const state = filters[groupName] || (filters[groupName] = { q: '', only: 'all' });

            const body = document.createElement('div');
            body.className = 'ov-gb';

            const tools = document.createElement('div');
            tools.className = 'ov-tools';
            const q = document.createElement('input');
            q.type = 'search';
            q.placeholder = 'Hľadať pracovisko…';
            q.value = state.q;
            const sel = document.createElement('select');
            sel.innerHTML = '<option value="all">Všetky</option><option value="run">Vyrába</option><option value="idle">Nevyrába</option>';
            sel.value = state.only;
            tools.appendChild(q);
            tools.appendChild(sel);

            const chips = document.createElement('div');
            chips.className = 'ov-chips';

            function paint() {
                chips.innerHTML = '';
                const needle = normalize(state.q);
                const shown = items.filter((w) => {
                    const producing = (w.count || 0) > 0;
                    if (state.only === 'run' && !producing) return false;
                    if (state.only === 'idle' && producing) return false;
                    if (!needle) return true;
                    return normalize(w.workcenter + ' ' + w.workcenterDescription).indexOf(needle) !== -1;
                });
                if (shown.length === 0) {
                    const empty = document.createElement('div');
                    empty.className = 'ov-empty';
                    empty.textContent = 'Nič nezodpovedá zadaniu.';
                    chips.appendChild(empty);
                    return;
                }
                shown.forEach((w) => chips.appendChild(chip(w)));
            }

            q.addEventListener('input', () => { state.q = q.value; paint(); });
            q.addEventListener('click', (e) => e.stopPropagation());
            sel.addEventListener('change', () => { state.only = sel.value; paint(); });

            body.appendChild(tools);
            body.appendChild(chips);
            paint();
            return body;
        }

        function render(host, list) {
            host.innerHTML = '';

            const hi = document.createElement('div');
            hi.className = 'ov-hi';
            hi.innerHTML = '<div class="k">VITAJTE SPÄŤ</div><div class="n"></div>' +
                '<div class="s">Vyberte pracovisko a začnite pracovať.</div>';
            hi.querySelector('.n').textContent = userName() || 'Pracoviská';
            host.appendChild(hi);

            // zoskupenie podla pola `area`; co ho nema, ide do "Ostatne"
            const groups = {};
            list.forEach((w) => {
                const key = w.area || 'Ostatné';
                (groups[key] = groups[key] || []).push(w);
            });
            const order = ORDER.filter((g) => groups[g])
                .concat(Object.keys(groups).filter((g) => ORDER.indexOf(g) === -1).sort());

            order.forEach((gName) => {
                const items = groups[gName].slice().sort((a, b) => String(a.workcenter).localeCompare(String(b.workcenter)));
                const online = items.reduce((acc, w) => acc + (w.count || 0), 0);
                const isOpen = expanded.has(gName);

                const box = document.createElement('div');
                box.className = 'ov-g' + (isOpen ? ' open' : '');

                const head = document.createElement('button');
                head.type = 'button';
                head.className = 'ov-gh';
                head.innerHTML =
                    '<div class="ov-ic">' + (ICONS[gName] || '🏭') + '</div>' +
                    '<div class="ov-gt"><b></b><span></span></div>' +
                    '<div class="ov-num"><b>' + items.length + '</b><span>PRACOVÍSK</span></div>' +
                    '<div class="ov-num on"><b>' + online + '</b><span>ONLINE</span></div>' +
                    '<div class="ov-ch">' + (isOpen ? '⌄' : '›') + '</div>';
                head.querySelector('.ov-gt b').textContent = gName;
                head.querySelector('.ov-gt span').textContent = SUBTITLES[gName] || '';
                head.addEventListener('click', () => {
                    if (expanded.has(gName)) expanded.delete(gName); else expanded.add(gName);
                    lastSignature = null;
                    ensureOverview();
                });
                box.appendChild(head);
                if (isOpen) box.appendChild(groupBody(gName, items));
                host.appendChild(box);
            });
        }

        function ensureOverview() {
            if (!document.getElementById(HOME_BUTTON_ID)) return;   // len uvodna obrazovka

            const panel = document.getElementById(PANEL_ID);
            const content = panel && panel.querySelector('.sapMPanelContent');
            if (!content) return;

            const list = workcenterList();
            if (!list.length) return;   // appka este taha data zo SAP (~25 s)

            injectStyles();
            document.body.classList.add(BODY_CLASS);

            let host = document.getElementById(OVERVIEW_ID);
            if (!host) {
                host = document.createElement('div');
                host.id = OVERVIEW_ID;
                content.insertBefore(host, content.firstChild);
            } else if (host.parentElement !== content) {
                content.insertBefore(host, content.firstChild);
            }

            // prekreslit len ked sa nieco zmenilo (clovek, pocty, rozbalenie)
            const signature = userName() + '#' +
                list.map((w) => w.workcenter + ':' + (w.count || 0)).join('|') + '#' +
                Array.from(expanded).join(',');
            if (signature === lastSignature) return;
            lastSignature = signature;

            render(host, list);
        }

        DomWatch.add(ensureOverview);
        // `count` sa meni bez zmeny DOM, preto aj casovac (rovnako ako Python verzia)
        onReady(() => setInterval(ensureOverview, 1500));
    }

    /* ---------- 3.9 Pracovny zoznam: zakazky ako kompaktne pilulky ---------- */

    /*
     * Lavy zoznam zakaziek v detaile pracoviska mal kazdu zakazku na styroch
     * riadkoch (~101 px) a okienko pevnu vysku 20em, takze na obrazovke boli
     * vidiet 2-3 zakazky a pod zoznamom ostavalo prazdne miesto az po spodok
     * stranky. Modul robi dve veci:
     *   1. kazdu polozku prekresli na kompaktnu "pilulku" (dva riadky, ~60 px)
     *   2. okienko so zoznamom natiahne az k spodnemu okraju obrazovky
     *
     * Povodny obsah polozky sa iba SKRYJE (CSS), v DOM ostava - klik, oznacenie,
     * prepinac VSET aj vyhladavanie appky funguju dalej bez zmeny. Udaje sa
     * citaju z modelu appky (binding context polozky), nie z textu na obrazovke,
     * takze sa nic neparsuje a nic sa nemoze "netrafit".
     */
    function modOrderListPills() {
        const LIST_ID = 'WorkcenterDetail--Work_List';
        const SCROLL_ID = 'WorkcenterDetail--WorkList_ScrollContainer';
        const LEFT_ID = 'WorkcenterDetail--LeftColumn_FlexBox';
        const STYLE_ID = '__pda_orderlist_styles__';
        const TIP_ID = '__pda_pill_tip__';
        const MIN_HEIGHT = 240;       // pod tuto vysku zoznam nikdy nestlacime
        const BOTTOM_GAP = 18;        // medzera pod lavym stlpcom

        let missingLeftLogged = false;
        let poslednyPodpis = '';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
#${LEFT_ID} { background:#fff !important; border:1px solid #dfe4ec !important; border-radius:14px !important;
  padding:8px !important; box-sizing:border-box; }
#${SCROLL_ID} { border:0 !important; background:transparent !important; }
#${LIST_ID} { background:transparent !important; }
/* kazdy zaznam = samostatna pilulka s plnym, jemnym ale viditelnym ramom;
   pozadie sa strieda (biela / svetlomodra), aby bolo vidiet, kde jedna konci */
#${LIST_ID} .sapMLIB.pda-pill-on { min-height:0 !important; height:auto !important; padding:0 !important;
  margin:4px 3px !important; border:1px solid #c3cfe0 !important; border-radius:10px !important;
  background:#fff !important; overflow:hidden; transition:border-color .12s, box-shadow .12s; }
#${LIST_ID} .sapMLIB.pda-pill-on:nth-child(even) { background:#eef3fa !important; }
#${LIST_ID} .sapMLIB.pda-pill-on:hover { border-color:#7ba4ee !important; box-shadow:0 3px 10px rgba(16,36,63,.12) !important; }
/* vybrana zakazka: tmavomodra na bielo - nedá sa zamenit so striedavym podfarbenim */
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected,
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected:nth-child(even) { border:2px solid #0b2447 !important;
  background:#13315c !important; box-shadow:0 3px 12px rgba(19,49,92,.35) !important; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .pda-pill { color:#fff; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .t-vyr { background:#fff; color:#13315c; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .t-zak { background:#2f5fa8; color:#fff; border-color:#5b87cc; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .mat { color:#fff; }
#${LIST_ID} .sapMLIB.pda-pill-on.sapMLIBSelected .matn { color:#c9d8ef; }
#${LIST_ID} .sapMLIB.pda-pill-on > *:not([data-pda-pill]) { display:none !important; }
.pda-pill { padding:5px 9px; font:12px/1.25 -apple-system,"Segoe UI",Roboto,sans-serif; color:#1a2233; }
.pda-pill.run { box-shadow: inset 4px 0 0 #2e9e4f; }
.pda-pill .r1 { display:flex; align-items:center; gap:5px; flex-wrap:nowrap; overflow:hidden; }
.pda-pill .r2 { display:flex; align-items:baseline; gap:5px; margin-top:2px; overflow:hidden; white-space:nowrap; }
.pda-pill .t { display:inline-block; padding:1px 7px; border-radius:999px; font-size:11.5px; font-weight:700;
  white-space:nowrap; letter-spacing:0; line-height:1.35; }
.pda-pill .t-vyr { background:#13315c; color:#fff; }
.pda-pill .t-zak { background:#dfeafc; color:#1b4f9c; border:1px solid #b9cdee; }
.pda-pill .t-run { background:#e7f6ec; color:#1d7a3c; border:1px solid #b6e2c5; margin-left:auto; }
.pda-pill .mat { font-size:11.5px; font-weight:700; color:#0f172a; white-space:nowrap; }
.pda-pill .matn { font-size:11.5px; color:#4a5568; overflow:hidden; text-overflow:ellipsis; white-space:nowrap; flex:1 1 auto; min-width:0; }
.pda-pill .t-vyr { flex:0 1 auto; min-width:0; overflow:hidden; text-overflow:ellipsis; }
.pda-pill .t-zak { flex:0 0 auto; }
/* bublina s celym obsahom zakazky */
#${TIP_ID} { display:none; position:fixed; z-index:100001; pointer-events:none;
  background:#fff; border:1px solid #b9cbe8; border-radius:10px; padding:10px 12px;
  box-shadow:0 10px 30px rgba(16,36,63,.22); max-width:520px;
  font:12.5px/1.45 -apple-system,"Segoe UI",Roboto,sans-serif; color:#1a2233; }
#${TIP_ID} .tab { display:grid; grid-template-columns:auto 1fr; gap:3px 12px; }
#${TIP_ID} .r { display:contents; }
#${TIP_ID} .k { color:#6b7c95; font-size:11px; text-transform:uppercase; letter-spacing:.05em;
  white-space:nowrap; padding-top:1px; }
#${TIP_ID} .v { color:#13315c; font-weight:600; }
`;
            document.head.appendChild(st);
        }

        /*
         * Udaje polozky. Najprv z modelu appky - pozor, zoznam moze byt naviazany
         * na POMENOVANY model, vtedy getBindingContext() bez mena vrati nic,
         * preto sa prejdu vsetky kontexty polozky (c.oBindingContexts).
         * Ked sa model neda precitat (v1.13.0 sa to tak aj stalo), pilulka sa
         * poskladá z textu, ktory appka do polozky napisala - tvar tych styroch
         * riadkov je pevny, takze sa to da bezpecne rozobrat.
         */
        function itemData(li) {
            const c = resolveControl(li);
            if (c) {
                const o = zKontextu(c);
                if (o) return o;
            }
            return zTextu(li);
        }

        function pouzitelne(o) {
            return o && (o.productionOrderNo || o.salesOrderNo || o.materialNo) ? o : null;
        }

        function zKontextu(c) {
            try {
                if (typeof c.getBindingContext === 'function') {
                    const ctx = c.getBindingContext();
                    const o = ctx && typeof ctx.getObject === 'function' ? ctx.getObject() : null;
                    if (pouzitelne(o)) return o;
                }
            } catch (e) { /* ignore */ }
            try {
                const mapa = c.oBindingContexts || {};
                for (const meno in mapa) {
                    const ctx = mapa[meno];
                    const o = ctx && typeof ctx.getObject === 'function' ? ctx.getObject() : null;
                    if (pouzitelne(o)) return o;
                }
            } catch (e) { /* ignore */ }
            return null;
        }

        // texty, ktore appka do polozky napisala (aj ked su uz skryte nasou pilulkou)
        const TEXT_SEL = '.sapMText, .sapMLabel, .sapMTitle, .sapMObjectIdentifierTitle,' +
                         '.sapMObjectIdentifierText, .sapMSLITitleOnly, .sapMSLIDescription';

        function riadkyTextu(li) {
            const out = [];
            li.querySelectorAll(TEXT_SEL).forEach((el) => {
                if (el.closest('[data-pda-pill]')) return;   // nase vlastne texty nie
                if (el.querySelector(TEXT_SEL)) return;      // len listy stromu
                const t = (el.textContent || '').replace(/\s+/g, ' ').trim();
                if (t && t !== '-' && out.indexOf(t) === -1) out.push(t);
            });
            return out;
        }

        /*
         * Tvar riadkov v polozke (overene na pracovisku 4836):
         *   7006003722 - 000010                        zakaznicka zakazka - polozka
         *   001600108392 - 0700 - 000000               vyrobna zakazka - operacia - sekvencia
         *   25288306 - S-PLATE CHARG.DOOR IM250E …     material - nazov
         *   WFT11CNC-B Horizontka 110 2 UP             popis operacie
         */
        function zTextu(li) {
            const d = {};
            riadkyTextu(li).forEach((t) => {
                let m;
                if ((m = t.match(/^(\d{6,12})\s*-\s*(\d{3,6})\s*-\s*(\d{3,6})$/))) {
                    d.productionOrderNo = m[1]; d.operationNo = m[2]; d.sequenceNo = m[3];
                } else if ((m = t.match(/^(\d{6,12})\s*-\s*(\d{3,6})$/))) {
                    d.salesOrderNo = m[1]; d.salesOrderItem = m[2];
                } else if ((m = t.match(/^(\d{6,10})\s*-\s*(\S.*)$/))) {
                    d.materialNo = m[1]; d.material = m[2];
                } else if (!d.description) {
                    d.description = t;
                }
            });
            return pouzitelne(d);
        }

        function txt(v) {
            return String(v === undefined || v === null ? '' : v).trim();
        }

        function span(cls, text) {
            const el = document.createElement('span');
            el.className = cls;
            el.textContent = text;
            return el;
        }

        /*
         * V pilulke su prve TRI riadky povodnej polozky (zakaznicka zakazka,
         * vyrobna zakazka s operaciou, material s nazvom). Stvrty riadok - popis
         * operacie - sa uz nezobrazuje, je len v bubline po nabehnuti mysou.
         */
        function buildPill(d) {
            const box = document.createElement('div');
            box.setAttribute('data-pda-pill', '1');
            box.className = 'pda-pill' + (Number(d.count) > 0 ? ' run' : '');

            // bez cisla zakazky nema zmysel ukazovat samotnu polozku ("000000")
            const zak = txt(d.salesOrderNo)
                ? [txt(d.salesOrderNo), txt(d.salesOrderItem)].filter(Boolean).join(' - ')
                : '';
            const vyroba = [txt(d.productionOrderNo), txt(d.operationNo), txt(d.sequenceNo)]
                .filter(Boolean).join(' - ');

            const r1 = document.createElement('div');
            r1.className = 'r1';
            // najprv vyrobna zakazka (tmavomodra), az potom zakaznicka (svetla)
            if (vyroba) r1.appendChild(span('t t-vyr', vyroba));
            if (zak) r1.appendChild(span('t t-zak', zak));
            if (Number(d.count) > 0) r1.appendChild(span('t t-run', '● vyrába'));
            box.appendChild(r1);

            const material = txt(d.materialNo).replace(/^0+/, '');
            const nazov = txt(d.material) || txt(d.descriptionShort);

            const r2 = document.createElement('div');
            r2.className = 'r2';
            if (material) r2.appendChild(span('mat', material));
            if (nazov) r2.appendChild(span('matn', nazov));
            if (r2.childNodes.length) box.appendChild(r2);

            udaje.set(box, d);   // pre bublinu pri nabehnuti mysou
            return box;
        }

        /* ---- bublina s celym obsahom zakazky (nabehnutie mysou) ----
         * Vlastna bublina namiesto systemoveho `title`: ukaze sa hned (nie po
         * sekunde), da sa ostylovat a zmestia sa do nej aj udaje, ktore sa do
         * pilulky nevojdu - hlavne stvrty riadok (popis operacie) a stav.
         */
        const udaje = new WeakMap();
        let tip = null;

        function tipEl() {
            if (tip && tip.isConnected) return tip;
            tip = document.createElement('div');
            tip.id = TIP_ID;
            document.body.appendChild(tip);
            return tip;
        }

        function tipRiadok(tabulka, popisok, hodnota) {
            if (!hodnota) return;
            const r = document.createElement('div');
            r.className = 'r';
            const k = document.createElement('span'); k.className = 'k'; k.textContent = popisok;
            const v = document.createElement('span'); v.className = 'v'; v.textContent = hodnota;
            r.appendChild(k); r.appendChild(v);
            tabulka.appendChild(r);
        }

        function ukazTip(pill, x, y) {
            const d = udaje.get(pill);
            if (!d) return;
            const el = tipEl();
            el.textContent = '';

            const t = document.createElement('div');
            t.className = 'tab';
            tipRiadok(t, 'Zákazka', [txt(d.salesOrderNo), txt(d.salesOrderItem)].filter(Boolean).join(' - '));
            tipRiadok(t, 'Výrobná zákazka', [txt(d.productionOrderNo), txt(d.operationNo), txt(d.sequenceNo)]
                .filter(Boolean).join(' - '));
            tipRiadok(t, 'Materiál', [txt(d.materialNo).replace(/^0+/, ''), txt(d.material) || txt(d.descriptionShort)]
                .filter(Boolean).join(' - '));
            // text operacie sa v bubline zamerne neukazuje (zelanie pouzivatela)
            tipRiadok(t, 'Stav', [txt(d.status), Number(d.count) > 0 ? 'vyrába (' + d.count + ')' : '']
                .filter(Boolean).join(' · '));
            el.appendChild(t);

            el.style.display = 'block';
            posunTip(x, y);
        }

        function posunTip(x, y) {
            if (!tip) return;
            const r = tip.getBoundingClientRect();
            const l = Math.min(Math.max(8, x + 16), W.innerWidth - r.width - 8);
            const t = y + 18 + r.height > W.innerHeight ? Math.max(8, y - r.height - 12) : y + 18;
            tip.style.left = Math.round(l) + 'px';
            tip.style.top = Math.round(t) + 'px';
        }

        function skryTip() {
            if (tip) tip.style.display = 'none';
        }

        // jedno odpocuvanie na celom zozname namiesto listenerov na kazdej pilulke
        function napojTip(list) {
            if (list.dataset.pdaTipOn) return;
            list.dataset.pdaTipOn = '1';
            list.addEventListener('mouseover', (e) => {
                const pill = e.target.closest && e.target.closest('[data-pda-pill]');
                if (pill) ukazTip(pill, e.clientX, e.clientY);
            });
            list.addEventListener('mousemove', (e) => {
                if (tip && tip.style.display === 'block') posunTip(e.clientX, e.clientY);
            });
            list.addEventListener('mouseleave', skryTip);
            list.addEventListener('mouseout', (e) => {
                const kam = e.relatedTarget;
                if (!kam || !kam.closest || !kam.closest('[data-pda-pill]')) skryTip();
            });
            // pri kliku (vybera sa zakazka) bublina prekaza
            list.addEventListener('click', skryTip);
        }

        /*
         * Vyska zoznamu: od jeho horneho okraja az po spodok obrazovky, minus to,
         * co je v lavom stlpci pod nim (graf "Resource over time"). Zvysok sa
         * pocita ako "cely lavy stlpec minus zoznam", takze nezavisi od toho,
         * ako vysoky zoznam prave je - nemoze sa to rozkmitat.
         */
        function fitHeight() {
            const sc = document.getElementById(SCROLL_ID);
            if (!sc || !sc.offsetParent) return;
            const left = document.getElementById(LEFT_ID);
            if (!left) {
                if (!missingLeftLogged) {
                    missingLeftLogged = true;
                    console.log(LOG, 'ľavý stĺpec detailu sa nenašiel, výšku zoznamu nemením');
                }
                return;
            }
            const top = sc.getBoundingClientRect().top;
            if (top <= 0) return;
            const podZoznamom = Math.max(0, left.scrollHeight - sc.offsetHeight);
            const ciel = Math.round(Math.max(MIN_HEIGHT, W.innerHeight - top - podZoznamom - BOTTOM_GAP));
            if (Math.abs(ciel - sc.offsetHeight) > 8) sc.style.height = ciel + 'px';
        }

        function apply() {
            const list = document.getElementById(LIST_ID);
            if (!list) return;
            injectStyles();

            napojTip(list);

            const polozky = list.querySelectorAll('.sapMLIB');
            let hotovych = 0;

            polozky.forEach((li) => {
                const d = itemData(li);
                if (!d) return;
                hotovych++;
                const key = [d.salesOrderNo, d.salesOrderItem, d.productionOrderNo, d.operationNo,
                             d.sequenceNo, d.materialNo, d.count].join('|');
                if (li.dataset.pdaPillKey === key && li.querySelector('[data-pda-pill]')) return;

                const stary = li.querySelector('[data-pda-pill]');
                if (stary) stary.remove();
                li.appendChild(buildPill(d));
                li.dataset.pdaPillKey = key;
                li.classList.add('pda-pill-on');
            });

            // aby bolo v konzole hned vidiet, ci sa udaje polozky daju precitat
            const podpis = hotovych + '/' + polozky.length;
            if (polozky.length && podpis !== poslednyPodpis) {
                poslednyPodpis = podpis;
                console.log(LOG, 'pilulky v pracovnom zozname:', podpis);
            }

            fitHeight();
        }

        DomWatch.add(apply);
        W.addEventListener('resize', fitHeight);
        onReady(apply);
    }

    /* -------- 3.10 Popis operacie: velke tlacidlo + okno cez obrazovku -------- */

    /*
     * Popis operacie bol v detaile vysadzany drobnym pismom do uzkeho stlpca a
     * pri dlhom texte sa nedal precitat. Modul povodny blok skryje a na jeho
     * miesto da velke tlacidlo; po kliknuti sa popis ukaze cez celu obrazovku
     * vo velkom pisme. Zatvara sa klikom mimo okna, krizikom alebo Esc.
     *
     * Prvky appky (overene v prehliadaci uz v Python verzii):
     *   WorkcenterDetail--Description_SimpleForm     - cely blok "Popis:"
     *   WorkcenterDetail--Description_FormattedText  - samotny text
     */
    function modOperationDescription() {
        const FORM_ID = 'WorkcenterDetail--Description_SimpleForm';
        const TEXT_ID = 'WorkcenterDetail--Description_FormattedText';
        const BTN_ID = '__pda_opis_button__';
        const OVERLAY_ID = '__pda_opis_overlay__';
        const STYLE_ID = '__pda_opis_styles__';
        const PRAZDNY = 'k tejto operácii nie je popis';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
.pda-opis-skryty { display:none !important; }
#${BTN_ID} { display:flex; align-items:center; gap:14px; width:100%; box-sizing:border-box; margin:6px 0 4px;
  padding:12px 18px; border:2px solid #b9cbe8; border-radius:14px; background:#f4f8ff; cursor:pointer;
  text-align:left; font:14px/1.4 -apple-system,"Segoe UI",Roboto,sans-serif; color:#13315c;
  transition:border-color .12s, box-shadow .12s, background .12s; }
#${BTN_ID}:hover { border-color:#2563eb; background:#eaf2ff; box-shadow:0 4px 14px rgba(16,36,63,.12); }
#${BTN_ID} .ikona { font-size:30px; line-height:1; flex:0 0 auto; }
#${BTN_ID} .stred { flex:1 1 auto; min-width:0; }
#${BTN_ID} .nadpis { display:block; font-size:15px; font-weight:800; letter-spacing:.06em; text-transform:uppercase; }
#${BTN_ID} .ukazka { display:block; font-size:12.5px; color:#5b6b83; margin-top:2px;
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
#${BTN_ID} .sipka { flex:0 0 auto; font-size:13px; font-weight:700; color:#2563eb; white-space:nowrap; }
#${BTN_ID}.prazdny { border-color:#e2e6ec; background:#fafbfc; color:#8a93a3; cursor:default; }
#${BTN_ID}.prazdny .sipka { display:none; }
#${OVERLAY_ID} { position:fixed; inset:0; background:rgba(10,20,40,.55); z-index:100000;
  display:flex; align-items:center; justify-content:center; padding:24px; box-sizing:border-box; }
#${OVERLAY_ID} .karta { background:#fff; border-radius:16px; width:min(1150px,94vw); max-height:90vh;
  display:flex; flex-direction:column; overflow:hidden; box-shadow:0 24px 70px rgba(16,36,63,.4);
  font-family:-apple-system,"Segoe UI",Roboto,sans-serif; }
#${OVERLAY_ID} .hl { display:flex; align-items:center; gap:14px; padding:14px 20px; background:#13315c; color:#fff; }
#${OVERLAY_ID} .hl .n { font-size:16px; font-weight:800; letter-spacing:.08em; text-transform:uppercase; }
#${OVERLAY_ID} .hl .z { font-size:13px; opacity:.8; }
#${OVERLAY_ID} .hl .x { margin-left:auto; background:none; border:0; color:#fff; font-size:30px; line-height:1;
  cursor:pointer; padding:0 4px; }
#${OVERLAY_ID} .telo { padding:24px 30px 30px; overflow:auto; font-size:21px; line-height:1.62; color:#17202e;
  white-space:pre-wrap; word-break:break-word; }
`;
            document.head.appendChild(st);
        }

        function textEl() {
            return document.getElementById(TEXT_ID) || document.querySelector('[id$="Description_FormattedText"]');
        }

        function formEl() {
            const f = document.getElementById(FORM_ID);
            if (f) return f;
            const t = textEl();
            return t ? (t.closest('.sapUiForm') || t.parentElement) : null;
        }

        function popisText() {
            const t = textEl();
            if (!t) return '';
            return String(t.innerText || t.textContent || '').replace(/ /g, ' ').trim();
        }

        // cislo zakazky do hlavicky okna (kozmetika - ked sa neda precitat, nic sa nedeje)
        function zakazkaPopis() {
            try {
                const main = W.sap.ui.getCore().byId('Main');
                const op = main && main.getController().getGlobals().getVar('oSelectedWorkcenterOperation');
                if (!op || !op.productionOrderNo) return '';
                return [op.productionOrderNo, op.operationNo].filter(Boolean).join(' / ');
            } catch (e) { return ''; }
        }

        function openOverlay() {
            const text = popisText();
            if (!text) return;
            const stary = document.getElementById(OVERLAY_ID);
            if (stary) stary.remove();

            const overlay = document.createElement('div');
            overlay.id = OVERLAY_ID;

            const karta = document.createElement('div');
            karta.className = 'karta';

            const hl = document.createElement('div');
            hl.className = 'hl';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = 'Popis operácie';
            const z = document.createElement('span'); z.className = 'z'; z.textContent = zakazkaPopis();
            const x = document.createElement('button'); x.className = 'x'; x.type = 'button';
            x.textContent = '×'; x.setAttribute('aria-label', 'Zavrieť');
            hl.appendChild(n); hl.appendChild(z); hl.appendChild(x);

            const telo = document.createElement('div');
            telo.className = 'telo';
            telo.textContent = text;

            karta.appendChild(hl); karta.appendChild(telo);
            overlay.appendChild(karta);

            const zavri = () => { overlay.remove(); document.removeEventListener('keydown', naEsc); };
            const naEsc = (e) => { if (e.key === 'Escape') zavri(); };

            x.addEventListener('click', zavri);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) zavri(); });
            document.addEventListener('keydown', naEsc);

            document.body.appendChild(overlay);
        }

        function apply() {
            const form = formEl();
            if (!form || !form.parentElement) return;
            injectStyles();

            let btn = document.getElementById(BTN_ID);
            if (!btn) {
                btn = document.createElement('button');
                btn.id = BTN_ID;
                btn.type = 'button';
                const ikona = document.createElement('span'); ikona.className = 'ikona'; ikona.textContent = '📄';
                const stred = document.createElement('span'); stred.className = 'stred';
                const nadpis = document.createElement('span'); nadpis.className = 'nadpis';
                nadpis.textContent = 'Popis operácie';
                const ukazka = document.createElement('span'); ukazka.className = 'ukazka';
                stred.appendChild(nadpis); stred.appendChild(ukazka);
                const sipka = document.createElement('span'); sipka.className = 'sipka'; sipka.textContent = 'otvoriť ▸';
                btn.appendChild(ikona); btn.appendChild(stred); btn.appendChild(sipka);
                btn.addEventListener('click', openOverlay);
            }
            if (btn.parentElement !== form.parentElement || btn.nextElementSibling !== form) {
                form.parentElement.insertBefore(btn, form);
            }
            if (!form.classList.contains('pda-opis-skryty')) form.classList.add('pda-opis-skryty');

            const text = popisText();
            const ukazka = btn.querySelector('.ukazka');
            const skratene = text.length > 120 ? text.slice(0, 120) + '…' : text;
            if (ukazka && ukazka.textContent !== (skratene || PRAZDNY)) {
                ukazka.textContent = skratene || PRAZDNY;
            }
            btn.classList.toggle('prazdny', !text);
            btn.disabled = !text;
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* ------------- 3.11 Krajsi graf vytazenia (Resource over time) ------------- */

    /*
     * Graf pod pracovnym zoznamom kresli appka cez Chart.js 4.4.2 do platna
     * `ResourceDetails`, detail (lupa) do `DialogChart` v dialogu
     * `Popups--Dialog_Chart` - zistene zo zdrojakov appky (PDA_Chart.displayTimelineChart).
     *
     * Modul graf NEKRESLI odznova, iba prestavi hotovu instanciu, takze udaje
     * ostavaju presne tie, ktore poslala appka.
     *
     * ⚠️ Dolezite (zistene 2026-09-14 pri v1.16.0): v Chart.js v4 NESTACI menit
     * `chart.options` - pri prekresleni sa nastavenia beru z `chart.config.options`
     * (a hotove osi maju este vlastnu kopiu v `chart.scales.x.options`). Preto sa
     * ta ista uprava nanasa na VSETKY tieto miesta; inak sa zmeni len hrubka pasov
     * (tie su v `chart.data.datasets`) a popisky casu ostanu dlhe.
     */
    function modChartStyle() {
        const MALY = 'ResourceDetails';
        const VELKY = 'DialogChart';
        const DIALOG_ID = 'Popups--Dialog_Chart';
        const STYLE_ID = '__pda_chart_styles__';

        // maly graf pod zoznamom: co najnizsi, nech ostane miesto na zakazky
        const M_RIADOK = 30, M_OKRAJE = 30, M_MIN = 84, M_MAX = 220;
        // detail: siroky, ale nie na celu vysku obrazovky
        const V_RIADOK = 80, V_OKRAJE = 110, V_MIN = 240;

        let chybaKnizniceLogged = false;
        let detailOtvoreny = false;
        let zatvaranieNapojene = false;

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
#WorkcenterDetail--ChartFlexBox { padding-top:0 !important; margin-top:0 !important; }
canvas#${MALY} { max-width:100% !important; }
/* nadpis "Resource over time" zabral dva riadky - staci jemny jednoriadkovy popisok */
.pda-chart-nadpis { font-size:11px !important; font-weight:700 !important; letter-spacing:.1em !important;
  text-transform:uppercase !important; color:#8e9bb0 !important; white-space:nowrap !important;
  line-height:1.2 !important; margin:0 !important; padding:0 !important; }
.pda-chart-nadpis .sapMTitleInner, .pda-chart-nadpis bdi, .pda-chart-nadpis span { font-size:11px !important;
  font-weight:700 !important; color:#8e9bb0 !important; white-space:nowrap !important; }
/* detail grafu (lupa): siroky na celu obrazovku, vysoky len tolko, kolko treba */
#${DIALOG_ID} { width:96vw !important; max-width:96vw !important; left:2vw !important;
  height:auto !important; max-height:82vh !important; top:9vh !important; transform:none !important;
  border-radius:16px !important; box-shadow:0 24px 70px rgba(16,36,63,.4) !important; }
#${DIALOG_ID} .sapMDialogSection, #${DIALOG_ID} .sapMDialogScrollCont { height:auto !important;
  max-height:none !important; padding:6px 12px 10px !important; }
#${DIALOG_ID} canvas#${VELKY} { max-width:100% !important; }
`;
            document.head.appendChild(st);
        }

        function kniznica() {
            const C = W.Chart;
            if (C && (typeof C.getChart === 'function' || C.instances)) return C;
            if (!chybaKnizniceLogged) {
                chybaKnizniceLogged = true;
                console.log(LOG, 'Chart.js sa nenašiel, graf len zmenším cez CSS');
            }
            return null;
        }

        function instancia(C, canvas) {
            try {
                if (typeof C.getChart === 'function') {
                    const ch = C.getChart(canvas);
                    if (ch) return ch;
                }
            } catch (e) { /* ignore */ }
            try {
                const zoznam = C.instances || {};
                for (const k in zoznam) {
                    if (zoznam[k] && zoznam[k].canvas === canvas) return zoznam[k];
                }
            } catch (e) { /* ignore */ }
            return null;
        }

        function dve(n) {
            return (n < 10 ? '0' : '') + n;
        }

        /*
         * Z hocijakeho tvaru casu spravi "07:00". Popisok moze prist ako cislo
         * (timestamp), Date alebo text "2026-09-14 07:00:00" - preto tri cesty.
         */
        function hhmm(v) {
            if (v === null || v === undefined) return v;
            if (v instanceof Date) return dve(v.getHours()) + ':' + dve(v.getMinutes());
            if (typeof v === 'number') {
                const d = new Date(v);
                return isNaN(d.getTime()) ? v : dve(d.getHours()) + ':' + dve(d.getMinutes());
            }
            const s = String(v);
            const m = s.match(/(\d{1,2}):(\d{2})/);
            return m ? dve(Number(m[1])) + ':' + m[2] : s;
        }

        /*
         * Popisky casu su u tohto grafu v `chart.data.labels`. Ukazalo sa (v1.16.1),
         * ze zmeny v `options` sa pri prekresleni nie vzdy prejavia - appka si graf
         * sama aktualizuje - kym zmeny v `data` drzia spolahlivo. Preto sa sekundy
         * odrezavaju priamo v popiskoch.
         * Prisny vzor: prepise sa LEN retazec, ktory je cely cas ("06:10:00",
         * "2026-09-14 06:10:00"). Mena pracovnikov ("4829: Maros Minarik") ostanu.
         */
        const CAS_CELY = /^\s*(?:\d{4}-\d{2}-\d{2}[ T])?(\d{1,2}):(\d{2})(?::\d{2})?\s*$/;

        function skratPopisky(chart) {
            const l = chart.data && chart.data.labels;
            if (!Array.isArray(l)) return false;
            let zmena = false;
            for (let i = 0; i < l.length; i++) {
                if (typeof l[i] !== 'string') continue;
                const m = l[i].match(CAS_CELY);
                if (!m) continue;
                const novy = dve(Number(m[1])) + ':' + m[2];
                if (novy !== l[i]) { l[i] = novy; zmena = true; }
            }
            return zmena;
        }

        function riadkov(chart) {
            try {
                const l = chart.data && chart.data.labels;
                if (l && l.length) return l.length;
            } catch (e) { /* ignore */ }
            return 1;
        }

        function vyskaMaleho(chart) {
            return Math.min(M_MAX, Math.max(M_MIN, riadkov(chart) * M_RIADOK + M_OKRAJE));
        }

        function vyskaVelkeho(chart) {
            const strop = Math.round(W.innerHeight * 0.66);
            return Math.min(strop, Math.max(V_MIN, riadkov(chart) * V_RIADOK + V_OKRAJE));
        }

        // vsetky miesta, odkial Chart.js berie nastavenia (viac o tom v komentari hore)
        function vsetkyOptions(chart) {
            const out = [];
            const pridaj = (o) => { if (o && out.indexOf(o) === -1) out.push(o); };
            try { pridaj(chart.options); } catch (e) { /* ignore */ }
            try { pridaj(chart.config && chart.config.options); } catch (e) { /* ignore */ }
            try { pridaj(chart.config && chart.config._config && chart.config._config.options); } catch (e) { /* ignore */ }
            return out;
        }

        function osX(velky) {
            return {
                maxRotation: 0,
                minRotation: 0,
                autoSkip: true,
                maxTicksLimit: velky ? 20 : 9,
                color: velky ? '#5b6b83' : '#8e9bb0',
                font: { size: velky ? 13 : 11 },
                callback(value) {
                    let raw = value;
                    try {
                        if (typeof this.getLabelForValue === 'function') raw = this.getLabelForValue(value);
                    } catch (e) { /* ignore */ }
                    return hhmm(raw);
                },
            };
        }

        function nastav(o, velky) {
            o.maintainAspectRatio = false;
            o.responsive = true;
            o.animation = false;
            o.layout = Object.assign({}, o.layout, { padding: { top: 2, right: 8, bottom: 0, left: 2 } });

            const sc = o.scales = o.scales || {};
            const x = sc.x = sc.x || {};
            x.ticks = Object.assign({}, x.ticks, osX(velky));
            x.grid = Object.assign({}, x.grid, {
                color: 'rgba(120,140,170,.14)', drawBorder: false, tickLength: 4,
            });
            x.border = Object.assign({}, x.border, { display: false });
            if (x.time) {
                x.time.displayFormats = Object.assign({}, x.time.displayFormats, {
                    millisecond: 'HH:mm', second: 'HH:mm', minute: 'HH:mm', hour: 'HH:mm', day: 'HH:mm',
                });
                x.time.tooltipFormat = 'HH:mm';
            }

            const y = sc.y = sc.y || {};
            y.ticks = Object.assign({}, y.ticks, {
                color: '#5b6b83', font: { size: velky ? 14 : 11.5, weight: '600' },
            });
            y.grid = Object.assign({}, y.grid, { display: false, drawBorder: false });
            y.border = Object.assign({}, y.border, { display: false });

            const p = o.plugins = o.plugins || {};
            // nadpis "Workcenter reports from last 24h: ..." je zbytocny
            p.title = Object.assign({}, p.title, { display: false });
            p.subtitle = Object.assign({}, p.subtitle, { display: false });
            p.tooltip = Object.assign({}, p.tooltip, {
                backgroundColor: 'rgba(19,49,92,.94)',
                titleFont: { size: velky ? 14 : 12.5 },
                bodyFont: { size: velky ? 14 : 12.5 },
                padding: 10,
                cornerRadius: 8,
            });
            if (p.legend) {
                p.legend.labels = Object.assign({}, p.legend.labels, {
                    boxWidth: 12, boxHeight: 12, usePointStyle: true,
                    color: '#5b6b83', font: { size: velky ? 13 : 11.5 },
                });
            }
        }

        function upravChart(chart, canvas) {
            if (!chart) return false;
            const velky = canvas.id === VELKY;

            // pasy: zaoblene, v detaile hrubsie
            (chart.data && chart.data.datasets ? chart.data.datasets : []).forEach((ds) => {
                ds.borderRadius = velky ? 9 : 6;
                ds.borderSkipped = false;
                ds.barThickness = velky ? 34 : 22;
                ds.maxBarThickness = velky ? 44 : 26;
                ds.borderWidth = 0;
            });

            skratPopisky(chart);
            vsetkyOptions(chart).forEach((o) => { try { nastav(o, velky); } catch (e) { /* ignore */ } });

            // hotove osi maju vlastnu kopiu nastaveni - bez toho ostanu dlhe popisky
            try {
                if (chart.scales && chart.scales.x && chart.scales.x.options) {
                    chart.scales.x.options.ticks = Object.assign({}, chart.scales.x.options.ticks, osX(velky));
                }
            } catch (e) { /* ignore */ }

            const wrap = canvas.parentElement;
            if (wrap) {
                wrap.style.position = wrap.style.position || 'relative';
                wrap.style.height = (velky ? vyskaVelkeho(chart) : vyskaMaleho(chart)) + 'px';
            }

            try { chart.resize(); } catch (e) { /* ignore */ }
            try { chart.update(); } catch (e) { /* ignore */ }

            if (!chart.__pdaUpraveny) {
                chart.__pdaUpraveny = true;
                let typOsi = '?';
                try { typOsi = chart.scales && chart.scales.x ? chart.scales.x.type : 'nie je'; } catch (e) { /* ignore */ }
                const ukazka = Array.isArray(chart.data && chart.data.labels)
                    ? chart.data.labels.slice(0, 3) : '(bez labels)';
                console.log(LOG, 'graf upravený:', canvas.id, '· riadkov:', riadkov(chart),
                            '· os X:', typOsi, '· popisky:', ukazka);
            }
            return true;
        }

        // dvojriadkovy nadpis nad malym grafom stlacime na jeden jemny riadok
        function zmensiNadpis() {
            const left = document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            if (!left) return;
            left.querySelectorAll('.sapMTitle, .sapMLabel, .sapMText').forEach((el) => {
                if (el.classList.contains('pda-chart-nadpis')) return;
                const t = (el.textContent || '').trim().toLowerCase();
                if (t === 'resource over time' || t === 'zdroj v čase' || t === 'ressource über zeit') {
                    el.classList.add('pda-chart-nadpis');
                }
            });
        }

        function detailJeOtvoreny() {
            const dlg = document.getElementById(DIALOG_ID);
            return !!(dlg && dlg.getBoundingClientRect().width > 0);
        }

        // klik vedla okna detail zavrie (appka na to vlastne tlacidlo nema)
        function napojZatvaranie() {
            if (zatvaranieNapojene) return;
            zatvaranieNapojene = true;
            document.addEventListener('mousedown', (e) => {
                const dlg = document.getElementById(DIALOG_ID);
                if (!dlg || dlg.getBoundingClientRect().width === 0) return;
                if (dlg.contains(e.target)) return;
                try {
                    const ctrl = getControl(DIALOG_ID);
                    if (ctrl && typeof ctrl.close === 'function') ctrl.close();
                } catch (err) { /* ignore */ }
            }, true);
            document.addEventListener('keydown', (e) => {
                if (e.key !== 'Escape' || !detailJeOtvoreny()) return;
                try {
                    const ctrl = getControl(DIALOG_ID);
                    if (ctrl && typeof ctrl.close === 'function') ctrl.close();
                } catch (err) { /* ignore */ }
            });
        }

        /*
         * Pri otvoreni detailu treba grafu poslat `resize` - do novej velkosti
         * dialogu sa sam neprekresli (overene uz v Python verzii).
         */
        function dopasujDetail(C) {
            const otvoreny = detailJeOtvoreny();
            if (otvoreny === detailOtvoreny) return;
            detailOtvoreny = otvoreny;
            if (!otvoreny) return;

            setTimeout(() => {
                try {
                    const canvas = document.getElementById(VELKY);
                    const ch = C && canvas ? instancia(C, canvas) : null;
                    if (ch) upravChart(ch, canvas);
                    W.dispatchEvent(new Event('resize'));
                } catch (e) { /* ignore */ }
            }, 150);
        }

        function apply() {
            injectStyles();
            napojZatvaranie();
            zmensiNadpis();

            const C = kniznica();
            dopasujDetail(C);

            [MALY, VELKY].forEach((id) => {
                const canvas = document.getElementById(id);
                if (!canvas) return;
                if (!C) {
                    // zaloha bez kniznice: aspon nizsie platno, Chart.js sa prekresli sam
                    const wrap = canvas.parentElement;
                    if (wrap && id === MALY && wrap.style.height !== M_MIN + 'px') {
                        wrap.style.height = M_MIN + 'px';
                    }
                    return;
                }
                const ch = instancia(C, canvas);
                if (!ch) return;
                if (!ch.__pdaUpraveny) {
                    upravChart(ch, canvas);
                } else if (skratPopisky(ch)) {
                    // appka graf prekreslila s dlhymi popiskami - znovu ich skratime
                    try { ch.update('none'); } catch (e) { /* ignore */ }
                }
            });
        }

        DomWatch.add(apply);
        // graf vznika az po odpovedi zo servera, preto este pomaly tik
        setInterval(apply, 1200);
        onReady(apply);
    }

    /* ------------- 3.12 Krajsia tabulka stavov (mriezka pri grafe) ------------- */

    /*
     * Tlacidlo s mriezkou pri grafe (WorkcenterDetail--Status_Overview_Button)
     * otvara dialog `Popups--TableSelectDialog_Overview` - holu SAP tabulku
     * so zaznamami stavov. Modul ju nechava tak, ako je (data aj klikanie),
     * len ju prekresli:
     *   - dialog siroky, zaobleny, tmavomodra hlavicka
     *   - datum a cas namiesto "2026-09-14 07:00:00" ako "14.09." + "07:00"
     *   - stav ako farebna pilulka podla tych istych pravidiel, ako maju
     *     stavove tlacidla (nastavenia -> Tlacidla - farby)
     *   - riadky vyssie, striedavo podfarbene, bez tvrdych ciar
     *
     * Povodny text bunky sa pamata v `data-pda-orig`, takze po prekresleni
     * tabulky sa prevod spravi znova a nic sa nestrati.
     */
    function modStatusTable() {
        const DIALOG_SEL = '[id^="Popups--TableSelectDialog_Overview"]';
        const STYLE_ID = '__pda_stable_styles__';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
${DIALOG_SEL}.sapMDialog { border-radius:16px !important; overflow:hidden !important;
  width:min(1450px,96vw) !important; max-width:96vw !important; max-height:86vh !important;
  box-shadow:0 24px 70px rgba(16,36,63,.38) !important; }
${DIALOG_SEL} .sapMDialogTitle, ${DIALOG_SEL} .sapMIBar.sapMHeader-CTX, ${DIALOG_SEL} .sapMDialogTitleGroup {
  background:#13315c !important; color:#fff !important; }
${DIALOG_SEL} .sapMDialogTitle .sapMTitle, ${DIALOG_SEL} .sapMIBar.sapMHeader-CTX .sapMTitle {
  color:#fff !important; font-size:15px !important; font-weight:800 !important; letter-spacing:.06em; }
${DIALOG_SEL} .sapMListTblHeader { background:#f3f6fb !important; }
${DIALOG_SEL} .sapMListTblHeaderCell, ${DIALOG_SEL} .sapMListTblHeaderCell .sapMLabel {
  font-size:11px !important; font-weight:800 !important; letter-spacing:.08em !important;
  text-transform:uppercase !important; color:#6b7c95 !important; }
${DIALOG_SEL} .sapMListTbl, ${DIALOG_SEL} .sapMList { background:#fff !important; }
${DIALOG_SEL} .sapMListTblRow { border-bottom:1px solid #eef2f7 !important; }
${DIALOG_SEL} .sapMListTblRow:nth-child(even) { background:#f7f9fd !important; }
${DIALOG_SEL} .sapMListTblRow:hover { background:#eaf1fc !important; }
${DIALOG_SEL} .sapMListTblCell { padding-top:9px !important; padding-bottom:9px !important;
  font-size:13px !important; color:#1a2233 !important; vertical-align:middle !important; }
${DIALOG_SEL} .sapMListTblRow.sapMLIBSelected { background:#13315c !important; }
${DIALOG_SEL} .sapMListTblRow.sapMLIBSelected .sapMListTblCell,
${DIALOG_SEL} .sapMListTblRow.sapMLIBSelected .sapMText { color:#fff !important; }
${DIALOG_SEL} .sapMSF, ${DIALOG_SEL} .sapMSFB { border-radius:10px !important; }
/* datum a cas */
.pda-cas { display:inline-flex; align-items:baseline; gap:6px; white-space:nowrap; }
.pda-cas .d { font-size:11px; color:#8e9bb0; }
.pda-cas .c { font-size:13.5px; font-weight:700; color:#13315c; }
.sapMLIBSelected .pda-cas .d { color:#b9cbe8; }
.sapMLIBSelected .pda-cas .c { color:#fff; }
/* stav ako pilulka */
.pda-stav { display:inline-block; padding:3px 11px; border-radius:999px; font-size:12px; font-weight:700;
  line-height:1.4; white-space:nowrap; background:#eef2f7; color:#41506a; }
/* ziadne lamanie cisel a nazvov cez pol stlpca */
${DIALOG_SEL} .sapMListTblCell, ${DIALOG_SEL} .sapMListTblHeaderCell { white-space:nowrap !important; }
${DIALOG_SEL} .sapMListTblCell .sapMText, ${DIALOG_SEL} .sapMListTblCell .sapMLabel {
  white-space:nowrap !important; word-break:normal !important; }
/* stlpec s materialom je najsirsi - nazov sa ma zmestit na jeden riadok */
${DIALOG_SEL} .pda-col-material { min-width:330px !important; }
.pda-sap { display:inline-block; margin-right:8px; padding:1px 8px; border-radius:999px;
  background:#13315c; color:#fff; font-size:11.5px; font-weight:700; }
.sapMLIBSelected .pda-sap { background:#fff; color:#13315c; }
`;
            document.head.appendChild(st);
        }

        // rovnake pravidlo ako pri farebnych tlacidlach: vyhrava najdlhsi sediaci text
        function pravidloPreStav(text) {
            const t = String(text || '').toLowerCase();
            let best = null;
            for (const r of buttonRules()) {
                const rt = String(r.text || '').toLowerCase();
                if (rt && t.indexOf(rt) !== -1 && (!best || rt.length > String(best.text || '').length)) best = r;
            }
            return best;
        }

        // "2026-09-14 07:00:00" / "2026-09-14T07:00" -> { den:"14.09.", cas:"07:00" }
        function rozlozCas(s) {
            const m = String(s || '').trim()
                .match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::\d{2})?$/);
            if (!m) return null;
            return { den: m[3] + '.' + m[2] + '.', cas: m[4] + ':' + m[5] };
        }

        function jeStav(text) {
            const t = String(text || '').trim();
            return t.length > 2 && t.length < 60 && !/^\d/.test(t) && !!pravidloPreStav(t);
        }

        function prekresliBunku(el) {
            const raw = el.dataset.pdaOrig !== undefined ? el.dataset.pdaOrig : (el.textContent || '').trim();
            if (!raw) return;
            // uz prevedene a od vtedy sa nic nezmenilo
            if (el.dataset.pdaOrig !== undefined && el.dataset.pdaDone === '1' &&
                el.querySelector('.pda-cas, .pda-stav')) return;

            const cas = rozlozCas(raw);
            if (cas) {
                el.dataset.pdaOrig = raw;
                el.dataset.pdaDone = '1';
                el.textContent = '';
                const w = document.createElement('span');
                w.className = 'pda-cas';
                const d = document.createElement('span'); d.className = 'd'; d.textContent = cas.den;
                const c = document.createElement('span'); c.className = 'c'; c.textContent = cas.cas;
                w.appendChild(d); w.appendChild(c);
                el.appendChild(w);
                return;
            }

            if (jeStav(raw)) {
                const r = pravidloPreStav(raw);
                el.dataset.pdaOrig = raw;
                el.dataset.pdaDone = '1';
                el.textContent = '';
                const p = document.createElement('span');
                p.className = 'pda-stav';
                p.textContent = raw;
                if (r) { p.style.background = r.bg; p.style.color = r.fg; }
                el.appendChild(p);
            }
        }

        /*
         * Stlpec "Materiál" je v tabulke najsirsi obsah (nazov vyrobku), pri
         * povodnej sirke sa lamal na styri riadky. Index stlpca sa najde podla
         * textu v hlavicke, takze to funguje aj keby appka stlpce prehodila.
         */
        function oznacStlpecMaterialu(dialog) {
            const hlavicky = dialog.querySelectorAll('.sapMListTblHeaderCell');
            if (!hlavicky.length) return;
            let idx = -1;
            hlavicky.forEach((h, i) => {
                if (idx < 0 && /materi/i.test(h.textContent || '')) idx = i;
            });
            if (idx < 0) return;
            if (!hlavicky[idx].classList.contains('pda-col-material')) {
                hlavicky[idx].classList.add('pda-col-material');
            }
            dialog.querySelectorAll('.sapMListTblRow').forEach((r) => {
                const bunky = r.querySelectorAll('.sapMListTblCell');
                if (bunky.length !== hlavicky.length || !bunky[idx]) return;   // stlpce nesedia, radsej nic
                if (!bunky[idx].classList.contains('pda-col-material')) {
                    bunky[idx].classList.add('pda-col-material');
                }
                doplnSapCislo(r, bunky[idx]);
            });
        }

        /*
         * V stlpci "Materiál" je nazov vyrobku, nie SAP cislo. Ak ho riadok
         * v datach ma (pole s "material" v nazve a hodnotou ako 25286609),
         * predradi sa pred nazov ako tmava pilulka. Ked ho data nemaju,
         * nestane sa nic.
         */
        function doplnSapCislo(row, bunka) {
            if (bunka.querySelector('.pda-sap')) return;
            let cislo = '';
            try {
                const c = resolveControl(row);
                const ctx = c && typeof c.getBindingContext === 'function' ? c.getBindingContext() : null;
                let o = ctx && typeof ctx.getObject === 'function' ? ctx.getObject() : null;
                if (!o && c && c.oBindingContexts) {
                    for (const meno in c.oBindingContexts) {
                        const x = c.oBindingContexts[meno];
                        const v = x && typeof x.getObject === 'function' ? x.getObject() : null;
                        if (v) { o = v; break; }
                    }
                }
                if (o) {
                    for (const k in o) {
                        if (!/material/i.test(k)) continue;
                        const v = String(o[k] === undefined || o[k] === null ? '' : o[k]).trim().replace(/^0+/, '');
                        if (/^\d{6,10}$/.test(v)) { cislo = v; break; }
                    }
                }
            } catch (e) { /* ignore */ }
            if (!cislo) return;

            const ciel = bunka.querySelector('.sapMText, .sapMLabel') || bunka;
            const p = document.createElement('span');
            p.className = 'pda-sap';
            p.textContent = cislo;
            ciel.insertBefore(p, ciel.firstChild);
        }

        function apply() {
            const dialog = document.querySelector(DIALOG_SEL + '.sapMDialog');
            if (!dialog) return;
            injectStyles();
            dialog.querySelectorAll('.sapMListTblCell .sapMText, .sapMListTblCell .sapMLabel')
                .forEach((el) => {
                    if (el.querySelector('.sapMText, .sapMLabel')) return;   // len listy stromu
                    try { prekresliBunku(el); } catch (e) { /* kozmetika, nikdy nesmie zhodit dialog */ }
                });
            try { oznacStlpecMaterialu(dialog); } catch (e) { /* ignore */ }
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* ---------- 3.13 Hlavicka detailu: vsetko do jedneho riadku ---------- */

    /*
     * Hlavicka nad stavovymi tlacidlami zabrala skoro tretinu obrazovky:
     * tri riadky formulara (Zakaznicka zakazka / Material / Production Order)
     * boli od seba na 50 px, tlacidlo "Operation Complete" bolo vysoke na dva
     * riadky a okienko VYKRES malo vlastny riadok nad tym vsetkym.
     *
     * Modul nic nepresklada v appke - len:
     *   1. nase okienko VYKRES presunie do toho isteho riadku, kde uz je
     *      prepinac Machine a tlacidlo Operation Complete (poradie: prepinac,
     *      vykres, tlacidlo)
     *   2. formular stlaci do kompaktneho zaobleneho boxu (riadky tesne pod sebou)
     *   3. tlacidlo Operation Complete da na jeden riadok, nizsie a sirsie
     *
     * Prvky appky: WorkcenterDetail--Order_FlexBox (VBox) -> OrderHeader_FlexBox
     * (HBox: formular, Machine_Switch, Confirm_Button).
     */
    function modDetailHeader() {
        const HEADER_ID = 'WorkcenterDetail--OrderHeader_FlexBox';
        const CONFIRM_ID = 'WorkcenterDetail--Confirm_Button';
        const SWITCH_ID = 'WorkcenterDetail--Machine_Switch';
        const DRAWING_ID = '__pda_order_drawing_wrapper__';
        const STYLE_ID = '__pda_detail_header_styles__';
        const COL_ID = '__pda_detail_rightcol__';
        const MACHINE_ID = '__pda_detail_machine__';

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* cely riadok hlavicky: prepinac, vykres a tlacidlo vedla seba, zvisle na stred */
#${HEADER_ID} { align-items:center !important; gap:12px !important; padding:4px 8px 2px !important; }

/* formular (zakazka / material / production order) ako kompaktny box */
#${HEADER_ID} .sapUiForm { background:#f7f9fd !important; border:1px solid #e3e9f1 !important;
  border-radius:12px !important; padding:7px 12px !important; margin:0 !important; flex:1 1 auto; min-width:0; }
#${HEADER_ID} .sapUiForm .sapUiFormElement,
#${HEADER_ID} .sapUiForm .sapUiRespGridRow { margin:0 !important; padding:0 !important; }
#${HEADER_ID} .sapUiForm [class*="sapUiRespGridSpan"],
#${HEADER_ID} .sapUiForm [class*="sapUiRespGridHSpace"] { padding-top:1px !important; padding-bottom:1px !important;
  margin-top:0 !important; margin-bottom:0 !important; }
#${HEADER_ID} .sapUiForm .sapMLabel { font-size:11px !important; line-height:1.45 !important;
  color:#6b7c95 !important; text-transform:uppercase; letter-spacing:.04em; font-weight:700 !important; }
#${HEADER_ID} .sapUiForm .sapMText, #${HEADER_ID} .sapUiForm .sapMTextMaxLine {
  font-size:13px !important; line-height:1.45 !important; color:#13315c !important; font-weight:600 !important; }
#${HEADER_ID} .sapUiForm .sapUiFormTitle, #${HEADER_ID} .sapUiForm .sapUiFormTitleH5 { display:none !important; }

/* prepinac Machine */
#${SWITCH_ID} { flex:0 0 auto; margin:0 !important; }

/* okienko VYKRES presunute do riadku */
#${DRAWING_ID}.pda-v-riadku { width:auto !important; margin:0 !important; flex:0 0 auto; }
#${DRAWING_ID}.pda-v-riadku > button { margin-right:0 !important; }

/* pravy stlpec: zhora VYKRES, pod nim Operation Complete, pod tym prepinac */
#${COL_ID} { display:flex; flex-direction:column; align-items:stretch; gap:8px;
  flex:0 0 auto; margin-left:auto; padding-left:12px; }
#${COL_ID} .pda-machine { display:flex; align-items:center; justify-content:flex-end; gap:8px; }
#${COL_ID} .pda-machine .sapMLabel { font-size:12px !important; color:#5b6b83 !important; }

/* Operation Complete: na jeden riadok, nizsie a sirsie */
#${CONFIRM_ID} { height:auto !important; min-height:0 !important; max-height:none !important;
  width:auto !important; flex:0 0 auto; margin:0 !important; }
#${CONFIRM_ID} .sapMBtnInner { height:auto !important; min-height:0 !important;
  padding:10px 26px !important; white-space:nowrap !important; }
#${CONFIRM_ID} .sapMBtnContent, #${CONFIRM_ID} bdi { white-space:nowrap !important;
  font-size:14px !important; line-height:1.2 !important; }
`;
            document.head.appendChild(st);
        }

        /*
         * Vpravo v hlavicke drzime vlastny stlpec a do neho presuvame (zhora dole):
         *   okienko VYKRES  ->  tlacidlo Operation Complete  ->  prepinac Machine
         * Prepinac ma popisok "Machine" ako samostatny prvok pred sebou, preto sa
         * oba davaju do maleho riadku, nech drzia spolu.
         */
        function pravyStlpec(header) {
            let col = document.getElementById(COL_ID);
            if (!col) {
                col = document.createElement('div');
                col.id = COL_ID;
            }
            if (col.parentElement !== header) header.appendChild(col);
            return col;
        }

        function riadokPrepinaca(col) {
            let row = document.getElementById(MACHINE_ID);
            if (!row) {
                row = document.createElement('div');
                row.id = MACHINE_ID;
                row.className = 'pda-machine';
            }
            if (row.parentElement !== col) col.appendChild(row);

            const sw = document.getElementById(SWITCH_ID);
            if (!sw) return;
            const popisok = sw.previousElementSibling &&
                            /machine/i.test(sw.previousElementSibling.textContent || '')
                ? sw.previousElementSibling : null;
            if (popisok && popisok.parentElement !== row) row.appendChild(popisok);
            if (sw.parentElement !== row) row.appendChild(sw);
        }

        function apply() {
            const header = document.getElementById(HEADER_ID);
            if (!header) return;
            injectStyles();

            const col = pravyStlpec(header);

            // 1) VYKRES uplne hore, cely vidno
            const wrap = document.getElementById(DRAWING_ID);
            if (wrap) {
                if (wrap.parentElement !== col) col.appendChild(wrap);
                if (!wrap.classList.contains('pda-v-riadku')) wrap.classList.add('pda-v-riadku');
            }
            // 2) tlacidlo Operation Complete
            const confirm = document.getElementById(CONFIRM_ID);
            if (confirm && confirm.parentElement !== col) col.appendChild(confirm);
            // 3) prepinac Machine uplne dole
            riadokPrepinaca(col);
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* -------------------- 3.14 Ladiaci vypis ---------------------------- */

    function modDebugLog() {
        XhrBus.subscribe((ev) => {
            console.log('--- [PDA executeBO] ---');
            console.log('request:', ev.requestRaw);
            console.log('response:', ev.responseRaw);
            console.log('-----------------------');
        });
    }

    /* ========================================================================
     *  4. ZOZNAM MODULOV
     *     Nove vylepsenie = napisat funkciu vyssie a pridat sem jeden riadok.
     * ====================================================================== */

    const MODULES = [
        {
            id: 'header',
            name: 'Vylepšená hlavička',
            desc: 'Slovenské popisky pri ikonách, väčšie meno používateľa, zvýraznené odhlásenie.',
            def: true,
            run: modEnhancedHeader,
        },
        {
            id: 'statusButtons',
            name: 'Farebné tlačidlá',
            desc: 'Farby tlačidiel podľa pravidiel v nastaveniach (predvolene: výroba zelená, prestoj oranžový, chyba červená). Po zapnutí „Nastavovanie tlačidiel“ meníš farby pravým klikom.',
            def: true,
            run: modButtonColors,
        },
        {
            id: 'crossSearch',
            name: 'Vyhľadávanie zákaziek',
            desc: 'Vyhľadávací riadok, ktorý hľadá naprieč všetkými pracoviskami naraz.',
            def: true,
            run: modCrossSearch,
        },
        {
            id: 'preventBack',
            name: 'Blokovanie tlačidla Späť',
            desc: 'Zabráni nechcenému vypadnutiu z aplikácie cez tlačidlo Späť v prehliadači.',
            def: true,
            run: modPreventBack,
        },
        {
            id: 'usersPanel',
            name: 'Panel rýchleho prepínania používateľov',
            desc: 'Bočný panel s tlačidlami na prihlásenie. Používateľov zadáš nižšie v nastaveniach.',
            def: false,
            needs: 'Zdieľaný terminál',
            run: modUsersPanel,
        },
        {
            id: 'scanner',
            name: 'Čiarový skener a RFID karty',
            desc: 'Číta skener aj čítačku kariet. Karta prepne používateľa, číslo zákazky sa vloží do vyhľadávania.',
            def: false,
            needs: 'Hardvér',
            run: modInputListener,
        },
        {
            id: 'drawing',
            name: 'Tlačidlo výkresu',
            desc: 'Načíta Excel s výkresmi a k otvorenej zákazke ukáže číslo výkresu a revíziu.',
            def: false,
            needs: 'Firemná sieť',
            run: modDrawingButton,
        },
        {
            id: 'overview',
            name: 'Prehľad pracovísk na úvode',
            desc: 'Zbalí úvodnú obrazovku do kategórií (Assembly, Welding, Machining…). Vidíš len čísla strojov a či pracujú; klik otvorí stroj.',
            def: true,
            run: modWorkcenterOverview,
        },
        {
            id: 'orderList',
            name: 'Zoznam zákaziek ako pilulky',
            desc: 'V detaile pracoviska stlačí každú zákazku do jedného kompaktného riadku a zoznam natiahne až po spodok obrazovky — zmestí sa ich viac.',
            def: true,
            run: modOrderListPills,
        },
        {
            id: 'opDescription',
            name: 'Popis operácie na celú obrazovku',
            desc: 'Namiesto drobného textu veľké tlačidlo „Popis operácie“; po kliknutí sa popis ukáže cez celú obrazovku vo veľkom písme (zavrie sa klikom vedľa alebo Esc).',
            def: true,
            run: modOperationDescription,
        },
        {
            id: 'chartStyle',
            name: 'Krajší graf vyťaženia',
            desc: 'Graf pod pracovným zoznamom: čas dole len ako HH:MM, zaoblené a tenšie pásy, nižšie plátno (uvoľní miesto zoznamu). Údaje sa nemenia, len vzhľad.',
            def: true,
            run: modChartStyle,
        },
        {
            id: 'statusTable',
            name: 'Krajšia tabuľka stavov',
            desc: 'Tabuľka, ktorá sa otvorí tlačidlom s mriežkou pri grafe: dátum a čas v krátkom tvare, stav ako farebná pilulka, vyššie a striedavo podfarbené riadky.',
            def: true,
            run: modStatusTable,
        },
        {
            id: 'detailHeader',
            name: 'Kompaktná hlavička detailu',
            desc: 'Zákazka, materiál a Production Order stlačí do jedného kompaktného boxu a prepínač Machine, okienko VÝKRES aj tlačidlo Operation Complete dá do jedného riadku. Uvoľní sa tým miesto dole.',
            def: true,
            run: modDetailHeader,
        },
        {
            id: 'debug',
            name: 'Ladiaci výpis do konzoly',
            desc: 'Vypisuje sieťovú komunikáciu aplikácie do konzoly prehliadača. Bežne netreba.',
            def: false,
            run: modDebugLog,
        },
    ];


    /* ------------------------------------------------------------------
     *  Subor s nastaveniami (zaloha / prenos medzi terminalmi).
     *  Prehliadac nevie pisat na disk podla cesty; vie to len do suboru, ktory
     *  clovek raz vyberie v systemovom okne "Ulozit ako" (File System Access API).
     *  Odkaz na subor sa uklada do IndexedDB, takze prezije obnovenie stranky;
     *  po restarte Chromu treba raz potvrdit pristup. Po kazdom ulozeni
     *  nastaveni sa subor prepise; "Nacitat zo suboru" nahra vsetko naspat
     *  (aj na inom pocitaci - napr. zo sietoveho disku).
     * ---------------------------------------------------------------- */
    const SETTINGS_FILE_KEYS = {
        modules: KEY_MODULES, users: KEY_USERS, pdm: KEY_PDM, excel: KEY_EXCEL,
        groups: KEY_GROUPS, buttons: KEY_BUTTONS, admin: KEY_ADMIN,
    };

    const SettingsFile = (function () {
        const DB = 'pda_settings_db', STORE = 'handles', KEY = 'settings_file_handle';
        let cached; // undefined = este necitane

        function openDb() {
            return new Promise((res, rej) => {
                const q = W.indexedDB.open(DB, 1);
                q.onupgradeneeded = () => q.result.createObjectStore(STORE);
                q.onsuccess = () => res(q.result);
                q.onerror = () => rej(q.error);
            });
        }
        async function get() {
            if (cached !== undefined) return cached;
            try {
                const db = await openDb();
                cached = await new Promise((res, rej) => {
                    const r = db.transaction(STORE, 'readonly').objectStore(STORE).get(KEY);
                    r.onsuccess = () => res(r.result || null);
                    r.onerror = () => rej(r.error);
                });
            } catch (e) { cached = null; }
            return cached;
        }
        async function set(handle) {
            const db = await openDb();
            await new Promise((res, rej) => {
                const tx = db.transaction(STORE, 'readwrite');
                tx.objectStore(STORE).put(handle, KEY);
                tx.oncomplete = res;
                tx.onerror = () => rej(tx.error);
            });
            cached = handle;
        }
        async function clear() {
            const db = await openDb();
            await new Promise((res, rej) => {
                const tx = db.transaction(STORE, 'readwrite');
                tx.objectStore(STORE).delete(KEY);
                tx.oncomplete = res;
                tx.onerror = () => rej(tx.error);
            });
            cached = null;
        }
        return { get, set, clear, supported: typeof W.showSaveFilePicker === 'function' };
    })();

    function suiteVersion() {
        try { return GM_info && GM_info.script ? String(GM_info.script.version) : ''; } catch (e) { return ''; }
    }

    // vsetko, co je v ulozisku Tampermonkey, v jednom objekte
    function collectSettings() {
        const out = {};
        for (const k of Object.keys(SETTINGS_FILE_KEYS)) out[k] = loadJson(SETTINGS_FILE_KEYS[k], null);
        return { aplikacia: 'PDA Suite', verzia: suiteVersion(), ulozene: new Date().toISOString(), nastavenia: out };
    }

    // nahra nastavenia zo suboru do uloziska; vrati pocet prevzatych casti
    function applySettingsObject(obj) {
        const n = obj && obj.nastavenia ? obj.nastavenia : obj;
        if (!n || typeof n !== 'object' || Array.isArray(n)) throw new Error('súbor neobsahuje nastavenia PDA Suite');
        let count = 0;
        for (const k of Object.keys(SETTINGS_FILE_KEYS)) {
            if (n[k] !== undefined && n[k] !== null) { saveJson(SETTINGS_FILE_KEYS[k], n[k]); count++; }
        }
        if (!count) throw new Error('v súbore nie je žiadna známa časť nastavení');
        return count;
    }

    async function ensureFilePermission(handle, mode, interactive) {
        let p = await handle.queryPermission({ mode });
        if (p === 'granted') return true;
        if (interactive) p = await handle.requestPermission({ mode });
        return p === 'granted';
    }

    async function writeSettingsToHandle(handle) {
        const w = await handle.createWritable();
        await w.write(JSON.stringify(collectSettings(), null, 2));
        await w.close();
    }

    // po ulozeni nastaveni prepise vybrany subor (ak je a je povoleny); nikdy nehadze chybu
    async function mirrorSettingsToFile(interactive) {
        try {
            const h = await SettingsFile.get();
            if (!h) return false;
            if (!(await ensureFilePermission(h, 'readwrite', !!interactive))) {
                console.log(LOG, 'súbor s nastaveniami: chýba povolenie na zápis (potvrď v nastaveniach)');
                return false;
            }
            await writeSettingsToHandle(h);
            console.log(LOG, 'nastavenia zapísané do súboru', h.name);
            return true;
        } catch (e) {
            console.warn(LOG, 'zápis nastavení do súboru zlyhal', e);
            return false;
        }
    }

    async function pickSettingsFile() {
        const h = await W.showSaveFilePicker({
            suggestedName: 'pda-suite-nastavenia.json',
            types: [{ description: 'Nastavenia PDA Suite', accept: { 'application/json': ['.json'] } }],
        });
        await SettingsFile.set(h);
        await writeSettingsToHandle(h);
        return h;
    }

    async function loadSettingsFromHandle(handle) {
        if (!(await ensureFilePermission(handle, 'read', true))) throw new Error('chýba povolenie na čítanie súboru');
        const file = await handle.getFile();
        return applySettingsObject(JSON.parse(await file.text()));
    }

    function downloadSettings() {
        const blob = new Blob([JSON.stringify(collectSettings(), null, 2)], { type: 'application/json' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'pda-suite-nastavenia.json';
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
    }

    /* ========================================================================
     *  5. PANEL NASTAVENI
     * ====================================================================== */

    const OVERLAY_ID = '__pda_settings_overlay__';
    const GEAR_ID = '__pda_settings_gear__';
    const PASS_ID = '__pda_settings_pass__';

    function injectSettingsStyles() {
        if (document.getElementById('__pda_settings_styles__')) return;
        const style = document.createElement('style');
        style.id = '__pda_settings_styles__';
        style.textContent = `
#${GEAR_ID} {
  position: fixed; right: 14px; bottom: 14px; z-index: 2147483000;
  width: 38px; height: 38px; border-radius: 50%; border: none;
  background: #313175; color: #fff; font-size: 19px; line-height: 1;
  cursor: pointer; opacity: .55; box-shadow: 0 2px 8px rgba(0,0,0,.3);
  transition: opacity .15s;
}
#${GEAR_ID}:hover { opacity: 1; }
#${OVERLAY_ID} {
  position: fixed; inset: 0; background: rgba(0,0,0,.5); z-index: 2147483001;
  display: flex; align-items: center; justify-content: center;
  font: 14px/1.5 -apple-system, "Segoe UI", Roboto, sans-serif; color: #1a1a1f;
}
.pda-set-box {
  background: #fff; border-radius: 12px; width: 94%; max-width: 720px;
  max-height: 88vh; display: flex; flex-direction: column; overflow: hidden;
  box-shadow: 0 10px 40px rgba(0,0,0,.35);
}
.pda-set-head {
  display: flex; justify-content: space-between; align-items: center;
  padding: 14px 18px; background: #313175; color: #fff; flex-shrink: 0;
}
.pda-set-head b { font-size: 1.02rem; }
.pda-set-x { background: none; border: 0; color: #fff; font-size: 24px; line-height: 1; cursor: pointer; }
.pda-set-body { padding: 6px 18px 18px; overflow: auto; }
.pda-set-body h3 {
  font-size: .76rem; text-transform: uppercase; letter-spacing: .07em;
  color: #6b7180; margin: 20px 0 8px; font-weight: 600;
}
.pda-mod {
  display: flex; gap: 11px; align-items: flex-start;
  padding: 10px 12px; border: 1px solid #e3e6eb; border-radius: 9px; margin-bottom: 7px;
}
.pda-mod input { margin-top: 3px; width: 17px; height: 17px; flex-shrink: 0; cursor: pointer; }
.pda-mod .nm { font-weight: 600; }
.pda-mod .ds { color: #5c6370; font-size: .87rem; }
.pda-badge {
  font-size: .66rem; text-transform: uppercase; letter-spacing: .05em;
  background: #eef0f4; color: #5c6370; border-radius: 20px; padding: 2px 8px; margin-left: 7px;
  white-space: nowrap;
}
.pda-set-body table { width: 100%; border-collapse: collapse; }
.pda-set-body th {
  text-align: left; font-size: .74rem; text-transform: uppercase;
  letter-spacing: .05em; color: #6b7180; padding: 0 6px 5px 0; font-weight: 600;
}
.pda-set-body td { padding: 0 6px 6px 0; }
.pda-set-body input[type=text], .pda-set-body input[type=password] {
  width: 100%; box-sizing: border-box; padding: 6px 9px;
  border: 1px solid #ccd1d9; border-radius: 6px; font: inherit; font-size: .9rem;
}
.pda-set-body input[type=color] { width: 38px; height: 30px; padding: 0; border: 1px solid #ccd1d9;
  border-radius: 6px; background: #fff; cursor: pointer; vertical-align: middle; }
.pda-set-body input[type=number] { width: 64px; padding: 6px 8px; border: 1px solid #ccd1d9; border-radius: 6px; font: inherit; font-size: .9rem; }
.pda-sw { display: inline-block; width: 16px; height: 16px; border-radius: 4px; border: 1px solid rgba(0,0,0,.15); vertical-align: middle; margin-right: 6px; }
.pda-del { background: #fdecea; border: 1px solid #f0b4ae; color: #b0201a;
  border-radius: 6px; cursor: pointer; padding: 5px 10px; font-size: .82rem; }
.pda-add { background: #eef0f4; border: 1px solid #ccd1d9; border-radius: 7px;
  cursor: pointer; padding: 6px 13px; font-size: .85rem; margin-top: 3px; }
.pda-note { color: #6b7180; font-size: .82rem; margin: 7px 0 0; }
.pda-set-foot {
  display: flex; justify-content: flex-end; gap: 9px; align-items: center;
  padding: 13px 18px; border-top: 1px solid #e3e6eb; flex-shrink: 0; background: #fafbfc;
}
.pda-btn-primary { background: #313175; color: #fff; border: 0; border-radius: 8px;
  padding: 9px 18px; font-weight: 600; cursor: pointer; font-size: .9rem; }
.pda-btn-plain { background: #fff; border: 1px solid #ccd1d9; border-radius: 8px;
  padding: 9px 16px; cursor: pointer; font-size: .9rem; }
`;
        document.head.appendChild(style);
    }

    function openSettings() {
        if (document.getElementById(OVERLAY_ID) || document.getElementById(PASS_ID)) return;
        const pwd = String((settings.admin && settings.admin.password) || '');
        if (!pwd) { openSettingsUnlocked(); return; }
        injectSettingsStyles();

        const ov = document.createElement('div');
        ov.id = PASS_ID;
        ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:2147483001;display:flex;' +
            'align-items:center;justify-content:center;font:14px/1.5 -apple-system,"Segoe UI",Roboto,sans-serif;color:#1a1a1f;';
        const box = document.createElement('div');
        box.style.cssText = 'background:#fff;border-radius:12px;width:92%;max-width:360px;padding:18px 20px;box-shadow:0 10px 40px rgba(0,0,0,.35);';
        const t = document.createElement('div');
        t.style.cssText = 'font-weight:600;font-size:1.02rem;margin-bottom:10px;';
        t.textContent = 'Nastavenia PDA Suite — heslo';
        const inp = document.createElement('input');
        inp.type = 'password';
        inp.autocomplete = 'off';
        inp.style.cssText = 'width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #ccd1d9;border-radius:8px;font:inherit;font-size:1rem;';
        const msg = document.createElement('div');
        msg.style.cssText = 'color:#b0201a;font-size:.85rem;min-height:1.2em;margin-top:6px;';
        const row = document.createElement('div');
        row.style.cssText = 'display:flex;justify-content:flex-end;gap:8px;margin-top:12px;';
        const cancel = document.createElement('button');
        cancel.type = 'button'; cancel.className = 'pda-btn-plain'; cancel.textContent = 'Zrušiť';
        const ok = document.createElement('button');
        ok.type = 'button'; ok.className = 'pda-btn-primary'; ok.textContent = 'Otvoriť';
        row.appendChild(cancel); row.appendChild(ok);
        box.appendChild(t); box.appendChild(inp); box.appendChild(msg); box.appendChild(row);
        ov.appendChild(box);
        document.body.appendChild(ov);
        setTimeout(() => inp.focus(), 50);

        const close = () => ov.remove();
        const submit = () => {
            if (inp.value === pwd) { close(); openSettingsUnlocked(); }
            else { msg.textContent = 'Nesprávne heslo.'; inp.value = ''; inp.focus(); }
        };
        cancel.addEventListener('click', close);
        ok.addEventListener('click', submit);
        inp.addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); if (e.key === 'Escape') close(); });
        ov.addEventListener('click', (e) => { if (e.target === ov) close(); });
    }

    function openSettingsUnlocked() {
        if (document.getElementById(OVERLAY_ID)) return;
        injectSettingsStyles();

        const draftModules = {};
        MODULES.forEach((m) => { draftModules[m.id] = isModuleOn(m); });
        const draftUsers = settings.users.map((u) => ({ ...u }));
        const draftPdm = { ...settings.pdm };
        const draftExcel = { ...settings.excel };
        const draftGroups = { value: String(settings.groups || '') };
        const draftButtons = buttonRules().map((r) => Object.assign({}, r));
        const draftAdmin = Object.assign({ password: '123456', pickMode: false }, settings.admin || {});

        const overlay = document.createElement('div');
        overlay.id = OVERLAY_ID;

        const box = document.createElement('div');
        box.className = 'pda-set-box';

        // --- hlavicka ---
        const head = document.createElement('div');
        head.className = 'pda-set-head';
        head.innerHTML = '<b>Nastavenia PDA Suite</b>';
        const closeBtn = document.createElement('button');
        closeBtn.className = 'pda-set-x';
        closeBtn.type = 'button';
        closeBtn.textContent = '×';
        head.appendChild(closeBtn);

        // --- telo ---
        const body = document.createElement('div');
        body.className = 'pda-set-body';

        const hMods = document.createElement('h3');
        hMods.textContent = 'Moduly';
        body.appendChild(hMods);

        MODULES.forEach((m) => {
            const row = document.createElement('div');
            row.className = 'pda-mod';

            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = draftModules[m.id];
            cb.addEventListener('change', () => { draftModules[m.id] = cb.checked; });

            const txt = document.createElement('div');
            const nm = document.createElement('div');
            nm.className = 'nm';
            nm.textContent = m.name;
            if (m.needs) {
                const badge = document.createElement('span');
                badge.className = 'pda-badge';
                badge.textContent = m.needs;
                nm.appendChild(badge);
            }
            const ds = document.createElement('div');
            ds.className = 'ds';
            ds.textContent = m.desc;
            txt.appendChild(nm);
            txt.appendChild(ds);

            const lbl = document.createElement('label');
            lbl.style.cssText = 'display:flex;gap:11px;align-items:flex-start;cursor:pointer;';
            lbl.appendChild(cb);
            lbl.appendChild(txt);
            row.appendChild(lbl);
            body.appendChild(row);
        });

        // --- pouzivatelia ---
        const hUsers = document.createElement('h3');
        hUsers.textContent = 'Používatelia pre rýchle prepínanie';
        body.appendChild(hUsers);

        const usersWrap = document.createElement('div');
        body.appendChild(usersWrap);

        function renderUsers() {
            usersWrap.innerHTML = '';
            const table = document.createElement('table');
            table.innerHTML =
                '<thead><tr><th>Meno (presne ako v zozname PDA)</th><th style="width:120px">Osobné číslo</th><th style="width:150px">ID karty</th><th style="width:40px"></th></tr></thead>';
            const tbody = document.createElement('tbody');

            draftUsers.forEach((u, i) => {
                const tr = document.createElement('tr');

                const mk = (field, type) => {
                    const td = document.createElement('td');
                    const inp = document.createElement('input');
                    inp.type = type;
                    inp.value = u[field] || '';
                    inp.addEventListener('input', () => { draftUsers[i][field] = inp.value.trim(); });
                    td.appendChild(inp);
                    return td;
                };

                tr.appendChild(mk('username', 'text'));
                tr.appendChild(mk('password', 'password'));
                tr.appendChild(mk('cardId', 'text'));

                const tdDel = document.createElement('td');
                const del = document.createElement('button');
                del.type = 'button';
                del.className = 'pda-del';
                del.textContent = '×';
                del.title = 'Odstrániť';
                del.addEventListener('click', () => { draftUsers.splice(i, 1); renderUsers(); });
                tdDel.appendChild(del);
                tr.appendChild(tdDel);

                tbody.appendChild(tr);
            });

            table.appendChild(tbody);
            usersWrap.appendChild(table);

            const add = document.createElement('button');
            add.type = 'button';
            add.className = 'pda-add';
            add.textContent = '+ Pridať používateľa';
            add.addEventListener('click', () => {
                draftUsers.push({ username: '', password: '', cardId: '' });
                renderUsers();
            });
            usersWrap.appendChild(add);

            const note = document.createElement('p');
            note.className = 'pda-note';
            note.textContent = 'Tieto údaje sa ukladajú len lokálne v Tampermonkey na tomto počítači — nikdy sa neposielajú na GitHub ani nikam inam.';
            usersWrap.appendChild(note);
        }
        renderUsers();

        // --- hromadne zadanie / prenos na iny pocitac ---
        const ioWrap = document.createElement('div');
        ioWrap.style.cssText = 'margin-top:12px;border:1px dashed #ccd1d9;border-radius:9px;padding:11px 12px;';

        const ioTitle = document.createElement('div');
        ioTitle.textContent = 'Hromadné zadanie / prenos na iný počítač';
        ioTitle.style.cssText = 'font-weight:600;font-size:.88rem;margin-bottom:4px;';

        const ioHelp = document.createElement('div');
        ioHelp.className = 'pda-note';
        ioHelp.style.marginTop = '0';
        const IO_HELP_DEFAULT = 'Jeden používateľ na riadok vo formáte:  Meno;osobné číslo;ID karty   (ID karty môže ostať prázdne)';
        ioHelp.textContent = IO_HELP_DEFAULT;

        const ta = document.createElement('textarea');
        ta.rows = 5;
        ta.spellcheck = false;
        ta.placeholder = 'Ján Novák;12345;0116984be8';
        ta.style.cssText = 'width:100%;box-sizing:border-box;margin-top:7px;padding:8px 10px;' +
            'border:1px solid #ccd1d9;border-radius:6px;font:12px/1.5 ui-monospace,Consolas,monospace;resize:vertical;';

        const ioBtns = document.createElement('div');
        ioBtns.style.cssText = 'display:flex;gap:8px;margin-top:7px;';

        const btnImport = document.createElement('button');
        btnImport.type = 'button';
        btnImport.className = 'pda-add';
        btnImport.style.marginTop = '0';
        btnImport.textContent = 'Načítať z textu';
        btnImport.addEventListener('click', () => {
            const parsed = ta.value
                .split('\n')
                .map((line) => line.trim())
                .filter((line) => line && line.indexOf(';') !== -1)
                .map((line) => {
                    const parts = line.split(';').map((s) => (s || '').trim());
                    return { username: parts[0], password: parts[1] || '', cardId: parts[2] || '' };
                })
                .filter((u) => u.username);

            if (parsed.length === 0) {
                ioHelp.style.color = '#b0201a';
                ioHelp.textContent = 'Nenašiel sa žiadny platný riadok. Formát je: Meno;osobné číslo;ID karty';
                return;
            }

            draftUsers.length = 0;
            parsed.forEach((u) => draftUsers.push(u));
            renderUsers();
            ioHelp.style.color = '';
            ioHelp.textContent = 'Načítaných používateľov: ' + parsed.length + '. Ešte to ulož tlačidlom dole.';
        });

        const btnExport = document.createElement('button');
        btnExport.type = 'button';
        btnExport.className = 'pda-add';
        btnExport.style.marginTop = '0';
        btnExport.textContent = 'Vypísať súčasných';
        btnExport.addEventListener('click', () => {
            ta.value = draftUsers.map((u) => [u.username, u.password, u.cardId].join(';')).join('\n');
            ioHelp.style.color = '';
            ioHelp.textContent = IO_HELP_DEFAULT;
        });

        ioBtns.appendChild(btnImport);
        ioBtns.appendChild(btnExport);
        ioWrap.appendChild(ioTitle);
        ioWrap.appendChild(ioHelp);
        ioWrap.appendChild(ta);
        ioWrap.appendChild(ioBtns);
        body.appendChild(ioWrap);

        // --- tlacidla: farby a nastavovanie pravym klikom ---
        const hBtn = document.createElement('h3');
        hBtn.textContent = 'Tlačidlá — farby';
        body.appendChild(hBtn);

        const pickRow = document.createElement('div');
        pickRow.className = 'pda-mod';
        const pickLbl = document.createElement('label');
        pickLbl.style.cssText = 'display:flex;gap:11px;align-items:flex-start;cursor:pointer;';
        const pickCb = document.createElement('input');
        pickCb.type = 'checkbox';
        pickCb.checked = !!draftAdmin.pickMode;
        pickCb.addEventListener('change', () => { draftAdmin.pickMode = pickCb.checked; });
        const pickTxt = document.createElement('div');
        pickTxt.innerHTML = '<div class="nm">Nastavovanie tlačidiel pravým klikom</div>' +
            '<div class="ds">Keď je zapnuté, pravý klik na ktorékoľvek tlačidlo v aplikácii ukáže paletu 12 farieb a „Reset tlačidla“. Výber sa hneď uloží. Po skončení to vypni, aby operátori omylom nemenili farby.</div>';
        pickLbl.appendChild(pickCb); pickLbl.appendChild(pickTxt);
        pickRow.appendChild(pickLbl);
        body.appendChild(pickRow);

        const btnWrap = document.createElement('div');
        body.appendChild(btnWrap);

        function renderButtons() {
            btnWrap.innerHTML = '';
            const table = document.createElement('table');
            table.innerHTML = '<thead><tr><th>Text tlačidla</th><th style="width:150px">ID (ak nemá text)</th><th style="width:70px">Poradie</th><th style="width:110px">Pozadie</th><th style="width:90px">Text</th><th style="width:40px"></th></tr></thead>';
            const tbody = document.createElement('tbody');
            if (draftButtons.length === 0) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 6; td.className = 'pda-note'; td.style.padding = '6px 0';
                td.textContent = 'Zatiaľ žiadne pravidlá. Zapni nastavovanie a klikni pravým na tlačidlo v aplikácii.';
                tr.appendChild(td); tbody.appendChild(tr);
            }
            draftButtons.forEach((r, i) => {
                const tr = document.createElement('tr');
                const mk = (field, type, extra) => {
                    const td = document.createElement('td');
                    const inp = document.createElement('input');
                    inp.type = type;
                    if (type === 'number') { inp.step = '0.5'; inp.value = Number.isFinite(Number(r[field])) ? r[field] : 1.5; }
                    else inp.value = r[field] || '';
                    if (extra) extra(inp, td);
                    inp.addEventListener('input', () => {
                        draftButtons[i][field] = type === 'number' ? Number(inp.value) : inp.value.trim();
                        if (field === 'bg') { draftButtons[i].fg = contrastColor(inp.value); renderButtons(); }
                    });
                    td.appendChild(inp);
                    return td;
                };
                tr.appendChild(mk('text', 'text'));
                tr.appendChild(mk('id', 'text'));
                tr.appendChild(mk('poradie', 'number'));
                tr.appendChild(mk('bg', 'color', (inp, td) => { const hex = document.createElement('span'); hex.style.cssText = 'font-size:.78rem;color:#6b7180;margin-left:6px;'; hex.textContent = r.bg || ''; td.appendChild(hex); }));
                tr.appendChild(mk('fg', 'color'));
                const tdDel = document.createElement('td');
                const del = document.createElement('button');
                del.type = 'button'; del.className = 'pda-del'; del.textContent = '×'; del.title = 'Odstrániť';
                del.addEventListener('click', () => { draftButtons.splice(i, 1); renderButtons(); });
                tdDel.appendChild(del);
                tr.appendChild(tdDel);
                tbody.appendChild(tr);
            });
            table.appendChild(tbody);
            btnWrap.appendChild(table);

            const tools = document.createElement('div');
            tools.style.cssText = 'display:flex;gap:8px;flex-wrap:wrap;margin-top:4px;';
            const add = document.createElement('button');
            add.type = 'button'; add.className = 'pda-add'; add.style.marginTop = '0'; add.textContent = '+ Pridať tlačidlo';
            add.addEventListener('click', () => { draftButtons.push({ text: '', id: '', bg: '#2563eb', fg: contrastColor('#2563eb'), poradie: 1.5 }); renderButtons(); });
            const defaults = document.createElement('button');
            defaults.type = 'button'; defaults.className = 'pda-add'; defaults.style.marginTop = '0'; defaults.textContent = 'Obnoviť predvolené farby';
            defaults.addEventListener('click', () => { draftButtons.length = 0; DEFAULT_BUTTON_RULES.forEach((r) => draftButtons.push(Object.assign({}, r))); renderButtons(); });
            tools.appendChild(add); tools.appendChild(defaults);
            btnWrap.appendChild(tools);

            const note = document.createElement('p');
            note.className = 'pda-note';
            note.textContent = 'Pravidlo sa hľadá podľa textu tlačidla (stačí časť textu), ID len keď tlačidlo text nemá. Poradie riadi zoradenie stavových tlačidiel na pracovisku (0 = prvé). Farba textu sa pri zmene pozadia dopočíta sama.';
            btnWrap.appendChild(note);
        }
        renderButtons();

        // --- vykresy ---
        const hPdm = document.createElement('h3');
        hPdm.textContent = 'Služba výkresov (PDM)';
        body.appendChild(hPdm);

        const pdmTable = document.createElement('table');
        pdmTable.innerHTML = '<thead><tr><th>Adresa služby</th><th style="width:210px">API kľúč (nepovinné)</th></tr></thead>';
        const pdmBody = document.createElement('tbody');
        const pdmTr = document.createElement('tr');

        const tdBase = document.createElement('td');
        const inpBase = document.createElement('input');
        inpBase.type = 'text';
        inpBase.value = draftPdm.base || '';
        inpBase.placeholder = 'http://172.16.77.134:9000';
        inpBase.addEventListener('input', () => { draftPdm.base = inpBase.value.trim(); });
        tdBase.appendChild(inpBase);

        const tdKey = document.createElement('td');
        const inpKey = document.createElement('input');
        inpKey.type = 'password';
        inpKey.value = draftPdm.key || '';
        inpKey.addEventListener('input', () => { draftPdm.key = inpKey.value.trim(); });
        tdKey.appendChild(inpKey);

        pdmTr.appendChild(tdBase);
        pdmTr.appendChild(tdKey);
        pdmBody.appendChild(pdmTr);
        pdmTable.appendChild(pdmBody);
        body.appendChild(pdmTable);

        // --- Excel s vykresmi ---
        const hExcel = document.createElement('h3');
        hExcel.textContent = 'Excel s výkresmi';
        body.appendChild(hExcel);

        const excelUrlTable = document.createElement('table');
        excelUrlTable.innerHTML = '<thead><tr><th>Adresa Excelu (nepovinné)</th></tr></thead>';
        const excelUrlBody = document.createElement('tbody');
        const excelUrlTr = document.createElement('tr');
        const tdUrl = document.createElement('td');
        const inpUrl = document.createElement('input');
        inpUrl.type = 'text';
        inpUrl.value = draftExcel.url || '';
        inpUrl.placeholder = 'C:\\Users\\meno\\OneDrive - HF MIXING GROUP\\HFSK O.4 Production - Data source\\AutomatedOQ180.xlsx';
        inpUrl.addEventListener('input', () => { draftExcel.url = inpUrl.value.trim(); });
        tdUrl.appendChild(inpUrl);
        excelUrlTr.appendChild(tdUrl);
        excelUrlBody.appendChild(excelUrlTr);
        excelUrlTable.appendChild(excelUrlBody);
        body.appendChild(excelUrlTable);

        const excelNote = document.createElement('p');
        excelNote.className = 'pda-note';
        excelNote.innerHTML =
            'Sem môžeš dať <b>cestu na disku</b> tak ako v Python appke (<code>C:\\…</code> alebo <code>\\\\server\\…</code>) — Excel sa načíta sám pri každom otvorení aplikácie. ' +
            '<b>Podmienka:</b> na stránke <code>chrome://extensions</code> → Tampermonkey → Podrobnosti zapnúť <b>„Povoliť prístup k URL adresám súborov"</b>.<br>' +
            'Funguje aj adresa <b>http(s)://</b>, ak by Excel visel na serveri.<br>' +
            'Ak pole necháš prázdne, súbor sa vyberá ručne tlačidlom <b>„Excel (nepovinné)"</b> a prehliadač si ho zapamätá.';
        body.appendChild(excelNote);

        const colTable = document.createElement('table');
        colTable.style.marginTop = '9px';
        colTable.innerHTML =
            '<thead><tr><th>Stĺpec — číslo zákazky</th><th>Stĺpec — číslo výkresu</th><th>Stĺpec — verzia</th></tr></thead>';
        const colBody = document.createElement('tbody');
        const colTr = document.createElement('tr');

        [['colOrder', 'H'], ['colDrawing', 'AH'], ['colVersion', 'AI']].forEach(([field, ph]) => {
            const td = document.createElement('td');
            const inp = document.createElement('input');
            inp.type = 'text';
            inp.value = draftExcel[field] || '';
            inp.placeholder = ph;
            inp.addEventListener('input', () => { draftExcel[field] = inp.value.trim().toUpperCase(); });
            td.appendChild(inp);
            colTr.appendChild(td);
        });

        colBody.appendChild(colTr);
        colTable.appendChild(colBody);
        body.appendChild(colTable);

        const colNote = document.createElement('p');
        colNote.className = 'pda-note';
        colNote.textContent = 'Písmená stĺpcov tak, ako ich vidíš v Exceli. Predvolene H, AH a AI. ' +
            'Excel je nepovinný — ak chýba, výkres sa hľadá priamo podľa čísla materiálu.';
        body.appendChild(colNote);


        // --- subor s nastaveniami ---
        const hFile = document.createElement('h3');
        hFile.textContent = 'Súbor s nastaveniami (záloha a prenos na iný počítač)';
        body.appendChild(hFile);

        const fileInfo = document.createElement('div');
        fileInfo.className = 'pda-note';
        fileInfo.style.marginTop = '0';
        fileInfo.textContent = 'Zisťujem…';
        body.appendChild(fileInfo);

        const fileMsg = document.createElement('div');
        fileMsg.className = 'pda-note';
        fileMsg.style.cssText = 'margin:4px 0 0;min-height:1.2em;color:#2f7d43;';

        const fileRow = document.createElement('div');
        fileRow.style.cssText = 'display:flex;gap:8px;flex-wrap:wrap;margin-top:8px;';
        const mkBtn = (label, title) => {
            const b = document.createElement('button');
            b.type = 'button'; b.className = 'pda-add'; b.style.marginTop = '0'; b.textContent = label;
            if (title) b.title = title;
            fileRow.appendChild(b);
            return b;
        };
        const bPick = mkBtn('Vybrať, kam sa má ukladať…', 'Systémové okno Uložiť ako — vyberieš miesto a názov súboru');
        const bSave = mkBtn('Uložiť do súboru teraz');
        const bLoad = mkBtn('Načítať zo súboru', 'Nahrá všetky nastavenia z vybraného súboru a obnoví stránku');
        const bDown = mkBtn('Stiahnuť kópiu', 'Uloží kópiu do priečinka Stiahnuté súbory');
        const bUp = mkBtn('Nahrať zo súboru…', 'Vyber .json súbor s nastaveniami');
        const bForget = mkBtn('Zabudnúť súbor');
        body.appendChild(fileRow);
        body.appendChild(fileMsg);

        const fileNote = document.createElement('p');
        fileNote.className = 'pda-note';
        fileNote.textContent = 'Po každom „Uložiť“ (aj po zmene farby pravým klikom) sa vybraný súbor prepíše. ' +
            'Na inom počítači: „Nahrať zo súboru…“, alebo vyber ten istý súbor na sieťovom disku a daj „Načítať“. ' +
            'Po reštarte Chromu môže prehliadač raz vyžiadať potvrdenie prístupu k súboru. ' +
            'Súbor obsahuje aj osobné čísla používateľov — ulož ho tam, kam majú prístup len vedúci.';
        body.appendChild(fileNote);

        const upInput = document.createElement('input');
        upInput.type = 'file'; upInput.accept = '.json,application/json'; upInput.style.display = 'none';
        body.appendChild(upInput);

        function fileSay(text, isError) { fileMsg.style.color = isError ? '#b0201a' : '#2f7d43'; fileMsg.textContent = text; }
        function fileRefresh() {
            SettingsFile.get().then((h) => {
                if (h) fileInfo.textContent = 'Súbor: ' + h.name + ' — prepisuje sa po každom uložení.';
                else fileInfo.textContent = SettingsFile.supported
                    ? 'Zatiaľ nie je vybraný žiadny súbor. Klikni „Vybrať, kam sa má ukladať…“.'
                    : 'Tento prehliadač nepodporuje výber miesta — použi „Stiahnuť kópiu“ a „Nahrať zo súboru…“.';
                bPick.disabled = bSave.disabled = !SettingsFile.supported;
                bForget.style.display = h ? '' : 'none';
            });
        }
        fileRefresh();

        bPick.addEventListener('click', () => {
            pickSettingsFile().then((h) => { fileSay('Uložené do ' + h.name + '.'); fileRefresh(); })
                .catch((e) => { if (!e || e.name !== 'AbortError') fileSay('Nepodarilo sa: ' + (e && e.message || e), true); });
        });
        bSave.addEventListener('click', () => {
            SettingsFile.get().then((h) => {
                if (!h) { fileSay('Najprv vyber, kam sa má ukladať.', true); return; }
                return mirrorSettingsToFile(true).then((ok) => fileSay(ok ? 'Zapísané do ' + h.name + '.' : 'Zápis zlyhal — povoľ prístup k súboru.', !ok));
            });
        });
        bLoad.addEventListener('click', () => {
            (async () => {
                let h = await SettingsFile.get();
                if (!h) {
                    if (typeof W.showOpenFilePicker !== 'function') { fileSay('Použi „Nahrať zo súboru…“.', true); return; }
                    const picked = await W.showOpenFilePicker({ types: [{ description: 'Nastavenia PDA Suite', accept: { 'application/json': ['.json'] } }], multiple: false });
                    h = picked[0];
                    await SettingsFile.set(h);
                }
                const n = await loadSettingsFromHandle(h);
                fileSay('Načítaných častí: ' + n + '. Obnovujem stránku…');
                shared.intentionalReload = true;
                setTimeout(() => location.reload(), 600);
            })().catch((e) => { if (!e || e.name !== 'AbortError') fileSay('Načítanie zlyhalo: ' + (e && e.message || e), true); });
        });
        bDown.addEventListener('click', () => { downloadSettings(); fileSay('Kópia sa sťahuje.'); });
        bUp.addEventListener('click', () => upInput.click());
        upInput.addEventListener('change', () => {
            const f = upInput.files && upInput.files[0];
            if (!f) return;
            f.text().then((txt) => {
                const n = applySettingsObject(JSON.parse(txt));
                fileSay('Nahraných častí: ' + n + '. Obnovujem stránku…');
                shared.intentionalReload = true;
                setTimeout(() => location.reload(), 600);
            }).catch((e) => fileSay('Nahratie zlyhalo: ' + (e && e.message || e), true));
        });
        bForget.addEventListener('click', () => { SettingsFile.clear().then(() => { fileSay('Súbor zabudnutý (na disku ostáva).'); fileRefresh(); }); });

        // --- heslo do nastaveni ---
        const hPwd = document.createElement('h3');
        hPwd.textContent = 'Heslo do nastavení';
        body.appendChild(hPwd);
        const pwdTable = document.createElement('table');
        pwdTable.innerHTML = '<thead><tr><th>Heslo (pýta sa pri otvorení ozubeného kolieska)</th></tr></thead>';
        const pwdBody = document.createElement('tbody');
        const pwdTr = document.createElement('tr');
        const pwdTd = document.createElement('td');
        const pwdInp = document.createElement('input');
        pwdInp.type = 'password';
        pwdInp.autocomplete = 'new-password';
        pwdInp.value = draftAdmin.password || '';
        pwdInp.addEventListener('input', () => { draftAdmin.password = pwdInp.value; });
        pwdTd.appendChild(pwdInp); pwdTr.appendChild(pwdTd); pwdBody.appendChild(pwdTr); pwdTable.appendChild(pwdBody);
        body.appendChild(pwdTable);
        const pwdNote = document.createElement('p');
        pwdNote.className = 'pda-note';
        pwdNote.textContent = 'Prázdne = bez hesla. Ukladá sa len lokálne v Tampermonkey na tomto počítači.';
        body.appendChild(pwdNote);

        // --- paticka ---
        const foot = document.createElement('div');
        foot.className = 'pda-set-foot';

        const cancel = document.createElement('button');
        cancel.type = 'button';
        cancel.className = 'pda-btn-plain';
        cancel.textContent = 'Zrušiť';

        const save = document.createElement('button');
        save.type = 'button';
        save.className = 'pda-btn-primary';
        save.textContent = 'Uložiť a obnoviť stránku';

        foot.appendChild(cancel);
        foot.appendChild(save);

        box.appendChild(head);
        box.appendChild(body);
        box.appendChild(foot);
        overlay.appendChild(box);
        document.body.appendChild(overlay);

        function close() {
            overlay.remove();
            document.removeEventListener('keydown', onEsc);
        }
        function onEsc(e) { if (e.key === 'Escape') close(); }

        closeBtn.addEventListener('click', close);
        cancel.addEventListener('click', close);
        overlay.addEventListener('click', (e) => { if (e.target === overlay) close(); });
        document.addEventListener('keydown', onEsc);

        save.addEventListener('click', () => {
            saveJson(KEY_MODULES, draftModules);
            saveJson(KEY_USERS, draftUsers.filter((u) => u.username));
            saveJson(KEY_PDM, draftPdm);
            saveJson(KEY_EXCEL, draftExcel);
            saveJson(KEY_GROUPS, draftGroups.value);
            saveJson(KEY_BUTTONS, draftButtons.map(normalizeRule).filter((r) => r.text || r.id));
            saveJson(KEY_ADMIN, { password: String(draftAdmin.password || ''), pickMode: !!draftAdmin.pickMode });
            close();
            mirrorSettingsToFile(true).then(() => {
                shared.intentionalReload = true;
                location.reload();
            });
        });
    }

    function ensureGear() {
        if (document.getElementById(GEAR_ID)) return;
        injectSettingsStyles();
        const gear = document.createElement('button');
        gear.id = GEAR_ID;
        gear.type = 'button';
        gear.title = 'Nastavenia PDA Suite';
        gear.textContent = '⚙';
        gear.addEventListener('click', openSettings);
        document.body.appendChild(gear);
    }

    /* ========================================================================
     *  6. SPUSTENIE
     * ====================================================================== */

    const active = [];
    MODULES.forEach((mod) => {
        if (!isModuleOn(mod)) return;
        try {
            mod.run();
            active.push(mod.id);
        } catch (e) {
            console.warn(LOG, 'modul "' + mod.id + '" sa nepodarilo spustiť', e);
        }
    });

    console.log(LOG, 'aktívne moduly:', active.length ? active.join(', ') : 'žiadne');

    onReady(ensureGear);
    DomWatch.add(ensureGear);

    try {
        GM_registerMenuCommand('Nastavenia PDA Suite', openSettings);
    } catch (e) { /* nedostupne mimo Tampermonkey */ }
})();
