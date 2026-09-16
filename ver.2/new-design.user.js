// ==UserScript==
// @name         PDA Suite NEW Design (HF Slovakia)
// @namespace    http://tampermonkey.net/pda-new-design
// @version      2.1.2
// @description  NOVY DIZAJN PDA - samostatna vetva vyvoja. Instaluje sa vedla povodneho skriptu, v Tampermonkey nechaj zapnuty vzdy len JEDEN z nich.
// @author       Gabris, Tvarozek
// @updateURL    https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/ver.2/new-design.user.js
// @downloadURL  https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/ver.2/new-design.user.js
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

    const LOG = '[PDA NEW Design]';
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
        pdmOpenDialog: null,     // doplni modul vykresu - okno so zoznamom vykresov
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
/* Plati pre VSETKY stavove tlacidla - aj pre osobny stav hore (Stretnutie,
   Prestavka, cakanie), lebo appka im dava tu istu triedu statusBtn.
   V riadku su rovnako vysoke: natiahnu sa na to najvyssie (dvojriadkove
   "Meranie v Procese s OTK"), text ostava zvisle na stred. */
.pda-status-row { align-content:flex-start !important; align-items:stretch !important; row-gap:0 !important; }
.statusBtn { height:auto !important; min-height:0 !important; margin:4px !important;
  align-self:stretch !important; border-radius:12px !important;
  border:2px solid #13315c !important;
  box-shadow:0 2px 6px rgba(16,36,63,.20) !important;
  transition:transform .13s ease, box-shadow .13s ease !important; }
.statusBtn .sapMBtnInner { height:100% !important; width:100% !important; min-height:0 !important;
  padding:11px 16px !important; border-radius:12px !important; box-shadow:none !important;
  display:flex !important; align-items:center !important; justify-content:center !important;
  box-sizing:border-box !important; }
/* sirka podla textu - nazov sa uz nezalomi ("Stretnuti / a", "Prestavk / a") */
.statusBtn { width:auto !important; min-width:130px !important; max-width:none !important; }
.statusBtn .sapMBtnContent, .statusBtn bdi { line-height:1.25 !important; white-space:nowrap !important;
  overflow:visible !important; text-overflow:clip !important; }
.statusBtn:hover { transform:translateY(-3px) !important;
  box-shadow:0 10px 20px rgba(16,36,63,.30) !important; }
.statusBtn:active { transform:translateY(-1px) !important;
  box-shadow:0 3px 8px rgba(16,36,63,.24) !important; }
/* panel "Osobny stav" na uvodnej obrazovke bol vysoky na pol obrazovky,
   hoci v nom su len tri tlacidla - zmensime jeho vnutorne odsadenie */
.pda-panel-tesny, .pda-panel-tesny .sapMPanelContent { min-height:0 !important; height:auto !important; }
.pda-panel-tesny .sapMPanelContent { padding-top:2px !important; padding-bottom:6px !important; }
.pda-panel-tesny .sapMPanelHdr, .pda-panel-tesny .sapMPanelHeaderTB {
  min-height:0 !important; padding-top:2px !important; padding-bottom:0 !important; }
.pda-panel-tesny .sapMFlexBox { min-height:0 !important; }
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
            // riadok, v ktorom stavove tlacidla sedia, musi natahovat na rovnaku vysku
            document.querySelectorAll('.statusBtn').forEach((b) => {
                const p = b.parentElement;
                if (p && !p.classList.contains('pda-status-row')) p.classList.add('pda-status-row');

                // panel osobneho stavu (uvodna obrazovka) je zbytocne vysoky;
                // panelov v detaile pracoviska sa nedotykame
                if (b.closest('[id^="WorkcenterDetail--"]')) return;
                const panel = b.closest('.sapMPanel');
                if (panel && !panel.classList.contains('pda-panel-tesny')) panel.classList.add('pda-panel-tesny');
            });
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

        // okno so zoznamom vykresov sprístupnime aj ostatnym modulom
        // (pouziva ho tlacidlo Vykresy v pravom paneli HF Slovakia)
        shared.pdmOpenDialog = pdmOpenDialog;

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
        const BOX1_ID = '__pda_left_box_zoznam__';
        const BOX2_ID = '__pda_left_box_graf__';
        const MIN_HEIGHT = 240;       // pod tuto vysku zoznam nikdy nestlacime
        const BOTTOM_GAP = 18;        // medzera pod lavym stlpcom

        let missingLeftLogged = false;
        let poslednyPodpis = '';
        let poslednyResize = 0;

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* Lavy stlpec: uzsi (pilulky maju sirku podla obsahu, siroky box nemal co
   vyplnit) a rozdeleny na dva samostatne zaoblene boxy - zoznam zakaziek
   a graf. Samotny stlpec uz teda ziadny ram nema. */
#${LEFT_ID} { background:transparent !important; border:0 !important; box-shadow:none !important;
  padding:0 !important; box-sizing:border-box;
  width:430px !important; max-width:430px !important; min-width:0 !important; flex:0 0 430px !important; }
.pda-left-box { background:#fff; border:1px solid #dfe4ec; border-radius:14px; padding:8px;
  margin-bottom:10px; box-sizing:border-box; box-shadow:0 1px 4px rgba(16,36,63,.06); }
.pda-left-box:last-child { margin-bottom:0; }
/* zaoblene hrany hore aj dole - vidno, kde posuvny zoznam konci */
#${SCROLL_ID} { border:0 !important; background:transparent !important; border-radius:14px !important; }
#${LIST_ID} { background:transparent !important; }
/* kazdy zaznam = samostatna pilulka s plnym, jemnym ale viditelnym ramom;
   pozadie sa strieda (biela / svetlomodra), aby bolo vidiet, kde jedna konci */
/* pilulka je siroka len tolko, kolko treba - vpravo uz neostava prazdny pas */
#${LIST_ID} .sapMLIB.pda-pill-on { min-height:0 !important; height:auto !important; padding:0 !important;
  margin:4px 3px !important; border:1px solid #c3cfe0 !important; border-radius:10px !important;
  background:#fff !important; overflow:hidden; transition:border-color .12s, box-shadow .12s;
  width:max-content !important; max-width:calc(100% - 10px) !important; }
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
            // ked je stlpec pripnuty na celu vysku (modul fullLeft), meriame po jeho
            // skutocny spodok, nie po spodok okna - inak by zoznam koncil privysoko
            const spodok = left.classList.contains('pda-layout-left')
                ? left.getBoundingClientRect().bottom
                : W.innerHeight;
            const ciel = Math.round(Math.max(MIN_HEIGHT, spodok - top - podZoznamom - BOTTOM_GAP));
            if (Math.abs(ciel - sc.offsetHeight) > 8) sc.style.height = ciel + 'px';
        }

        /*
         * Lavy stlpec sa rozdeli na dva zaoblene boxy: v prvom zoznam zakaziek,
         * v druhom graf. Povodne deti stlpca sa presunu do nasich obalov - to,
         * co obsahuje zoznam (a vsetko pred nim), ide do prveho, zvysok do druheho.
         * Ked appka stlpec prekresli, obaly zmiznu a pri dalsom tiku sa spravia
         * znova.
         */
        function rozdelStlpec() {
            const left = document.getElementById(LEFT_ID);
            const sc = document.getElementById(SCROLL_ID);
            if (!left || !sc) return;

            let b1 = document.getElementById(BOX1_ID);
            let b2 = document.getElementById(BOX2_ID);
            if (!b1) { b1 = document.createElement('div'); b1.id = BOX1_ID; b1.className = 'pda-left-box'; }
            if (!b2) { b2 = document.createElement('div'); b2.id = BOX2_ID; b2.className = 'pda-left-box'; }
            if (b1.parentElement !== left) left.insertBefore(b1, left.firstChild);
            if (b2.parentElement !== left) left.appendChild(b2);

            let zaZoznamom = false;
            let presunute = false;
            Array.from(left.children).forEach((ch) => {
                if (ch === b1 || ch === b2) return;
                if (ch.contains(sc)) {
                    if (ch.parentElement !== b1) { b1.appendChild(ch); presunute = true; }
                    zaZoznamom = true;
                    return;
                }
                const ciel = zaZoznamom ? b2 : b1;
                if (ch.parentElement !== ciel) { ciel.appendChild(ch); presunute = true; }
            });

            /*
             * Po presune ma graf novy ramec - Chart.js sa prepocita az po `resize`,
             * inak by ostal neviditelny (vysoky 0 px). `resize` vsak rozhybe cely
             * UI5, takze ho posielame najviac raz za sekundu; bez tejto poistky
             * by sa pri kazdom prekresleni appky mohla rozbehnut spatna vazba
             * (presun -> resize -> prekreslenie -> presun) a appka by zamrzla.
             */
            if (presunute && Date.now() - poslednyResize > 1000) {
                poslednyResize = Date.now();
                setTimeout(() => W.dispatchEvent(new Event('resize')), 60);
            }
        }

        function apply() {
            const list = document.getElementById(LIST_ID);
            if (!list) return;
            injectStyles();
            rozdelStlpec();

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
#${BTN_ID} { display:flex; align-items:center; gap:14px; width:100%; box-sizing:border-box; margin:6px 0 20px;
  padding:12px 18px; border:2px solid #b9cbe8; border-radius:14px; background:#f4f8ff; cursor:pointer;
  text-align:left; font:14px/1.4 -apple-system,"Segoe UI",Roboto,sans-serif; color:#13315c;
  box-shadow:0 2px 6px rgba(16,36,63,.10);
  transition:transform .13s ease, box-shadow .13s ease, border-color .13s, background .13s; }
/* rovnako "plasticke" ako stavove tlacidla - pri prechode mysou sa nadvihne */
#${BTN_ID}:hover { border-color:#2563eb; background:#eaf2ff; transform:translateY(-3px);
  box-shadow:0 10px 20px rgba(16,36,63,.22); }
#${BTN_ID}:active { transform:translateY(-1px); box-shadow:0 3px 8px rgba(16,36,63,.18); }
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
        const V_RIADOK = 80, V_OKRAJE = 165, V_MIN = 260;   // vacsie okraje = miesto na popisky casu

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
/* poistka: kym kod graf nedoladi, ma okno aspon nejaku vysku - inak by sa
   na okamih ukazalo prazdne (vysku platna nastavuje az JS) */
#${DIALOG_ID} .sapMDialogScrollCont { min-height:300px !important; }
#${DIALOG_ID} canvas#${VELKY} { max-width:100% !important; min-height:260px !important; }
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
            const strop = Math.round(W.innerHeight * 0.74);
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

            /*
             * Graf v okne appka vytvara az po otvoreni dialogu. Predtym sa tu
             * cakalo pevnych 150 ms a ked este nebol hotovy, doladil ho az
             * pomaly tik (1,2 s) - okno preto dlho vyzeralo prazdne. Teraz sa
             * skusa kazdych 60 ms, takze sa chyti hned, ako vznikne.
             * Globalny `resize` (prekresli vsetky grafy na stranke) ostava uz
             * len ako zaloha, ked sa instancia vobec nenajde.
             */
            let pokus = 0;
            const skus = () => {
                if (!detailJeOtvoreny()) return;
                try {
                    const canvas = document.getElementById(VELKY);
                    const ch = C && canvas ? instancia(C, canvas) : null;
                    if (ch) { upravChart(ch, canvas); return; }
                } catch (e) { /* ignore */ }
                if (++pokus < 40) setTimeout(skus, 60);
                else { try { W.dispatchEvent(new Event('resize')); } catch (e) { /* ignore */ } }
            };
            setTimeout(skus, 30);
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
        const STATUS_ID = 'WorkcenterDetail--Order_Status_Flexbox';
        const ORDER_ID = 'WorkcenterDetail--Order_FlexBox';

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
/* popisok aj hodnota musia sediet na tom istom riadku - rovnaka vyska riadku
   v px (nie nasobok), inak ich rozne velke pisma posunu voci sebe */
#${HEADER_ID} .sapUiForm .sapMLabel { font-size:11px !important; line-height:22px !important;
  color:#6b7c95 !important; text-transform:uppercase; letter-spacing:.04em; font-weight:700 !important;
  display:inline-block !important; vertical-align:middle !important; }
#${HEADER_ID} .sapUiForm .sapMText, #${HEADER_ID} .sapUiForm .sapMTextMaxLine {
  font-size:13px !important; line-height:22px !important; color:#13315c !important; font-weight:600 !important;
  display:inline-block !important; vertical-align:middle !important; }
#${HEADER_ID} .sapUiForm .sapMLabel .sapMLabelTextWrapper,
#${HEADER_ID} .sapUiForm .sapMText .sapMTextMaxLine { line-height:22px !important; }
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

/* Stavove tlacidla su hore nad boxom so zakazkou a sedia vo vlastnom farebnom
   banery. Zarovnanie s boxom ide cez margin-left a max-width (inline styl),
   takze vnutorne odsadenie tu uz pokojne byt moze. */
#${STATUS_ID} { background:linear-gradient(180deg,#eef3fa 0%,#dde7f4 100%) !important;
  border:1px solid #c9d7ea !important; border-radius:14px !important;
  padding:8px 10px !important;
  box-shadow:0 2px 8px rgba(16,36,63,.10) !important; margin-bottom:10px !important;
  box-sizing:border-box !important; }

/* bezici cinnost vpravo (Vyroba - Vyroba / meno / cas / Zastavit) ako jemna pilulka */
.pda-aktivita { border-radius:14px !important; background:#e9f6ed !important;
  border:1px solid #bfe3ca !important; box-shadow:0 2px 8px rgba(16,36,63,.10) !important;
  padding:12px 14px !important; margin:8px 0 !important; box-sizing:border-box; }
.pda-aktivita .sapMBtn { border-radius:999px !important; box-shadow:0 2px 6px rgba(16,36,63,.16) !important; }
.pda-aktivita .sapMBtn .sapMBtnInner { border-radius:999px !important; padding:6px 16px !important; }

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

        /*
         * Lavy okraj stavovych tlacidiel a pilulky POPIS OPERACIE zarovnany
         * s boxom hore (zakazka / material / production order). Odsadenie sa
         * nehada v pixeloch - odmeria sa priamo na stranke, takze to sedi aj
         * pri inej sirke okna. Zmena `style` nespusti DomWatch (sleduje len
         * pridavanie a mazanie prvkov), takze sa to nemoze rozkmitat.
         */
        function zarovnajVlavo() {
            const form = document.querySelector('#' + HEADER_ID + ' .sapUiForm');
            if (!form) return;
            const ciel = form.getBoundingClientRect().left;
            if (!ciel) return;

            /*
             * Baner s tlacidlami zarovnavame CELY (jeho vlastny lavy aj pravy
             * okraj) s boxom pod nim - nie prve tlacidlo v nom. Odkedy ma baner
             * viditelne pozadie, vycnieval by vlavo, aj keby tlacidlo sedelo.
             */
            const stav = document.getElementById(STATUS_ID);
            if (stav) {
                const sr = stav.getBoundingClientRect();
                posun(stav, 'marginLeft', ciel - sr.left);
                const sirkaBanera = Math.round(form.getBoundingClientRect().right - Math.max(sr.left, ciel));
                if (sirkaBanera > 200 && Math.abs(sirkaBanera - sr.width) > 2) {
                    stav.style.maxWidth = sirkaBanera + 'px';
                }
            }

            const opis = document.getElementById('__pda_opis_button__');
            if (!opis) return;
            posun(opis, 'marginLeft', ciel - opis.getBoundingClientRect().left);

            /*
             * Pilulka POPIS bola siroka na celu plochu a podliezala pravy stlpec
             * (Components / BOM, Parallel Process Handling). Stavove tlacidla
             * siahaju az popod ten stlpec (len su vyssie, takze sa nebiju), preto
             * nestaci zarovnat podla nich - hlada sa aj lavy okraj praveho stlpca
             * a berie sa to, co je viac vlavo.
             */
            const r = opis.getBoundingClientRect();
            let koniec = Infinity;

            const tlacidla = stav ? stav.querySelectorAll('.statusBtn') : null;
            if (tlacidla && tlacidla.length) {
                koniec = tlacidla[tlacidla.length - 1].getBoundingClientRect().right;
            }
            const stlpec = lavyOkrajPravehoStlpca(r.left);
            if (stlpec) koniec = Math.min(koniec, stlpec - 14);

            if (!isFinite(koniec)) return;
            const sirka = Math.round(koniec - r.left);
            if (sirka > 200 && Math.abs(sirka - r.width) > 2) opis.style.maxWidth = sirka + 'px';
        }

        /*
         * Zelene okienko s beziacou cinnostou vpravo (Vyroba - Vyroba, meno,
         * cas, Zastavit). Najde sa podla tlacidla "Zastavit" - jeho najblizsi
         * obal je to okienko.
         */
        function oznacAktivitu() {
            document.querySelectorAll('.sapMBtn').forEach((btn) => {
                const t = (btn.textContent || '').trim();
                if (!/^(zastavi|stop)/i.test(t)) return;
                const box = btn.closest('.sapMLIB') || btn.closest('.sapMVBox') ||
                            btn.closest('.sapMFlexBox') || btn.parentElement;
                if (box && !box.classList.contains('pda-aktivita')) box.classList.add('pda-aktivita');
            });
        }

        // lavy okraj praveho stlpca detailu (Components / BOM, Parallel Process Handling)
        function lavyOkrajPravehoStlpca(odX) {
            let x = Infinity;
            document.querySelectorAll('.sapMBtn, .sapMTitle, .sapMPanel').forEach((el) => {
                const t = (el.textContent || '').trim();
                if (!/^components\b|parallel process/i.test(t)) return;
                const er = el.getBoundingClientRect();
                if (er.width && er.left > odX + 150) x = Math.min(x, er.left);
            });
            return isFinite(x) ? x : null;
        }

        /*
         * Popisok a hodnota v hornom boxe nesedeli na jednej linke - popisok bol
         * vyssie. Presnu hodnotu odsadenia sa neda spolahlivo uhadnut (zavisi od
         * temy a velkosti pisma), preto sa rozdiel ich zvislych stredov odmeria
         * priamo na stranke a popisok sa o nho posunie. Po posune je rozdiel
         * nulovy, takze sa to samo ustali.
         */
        function zarovnajRiadky() {
            const form = document.querySelector('#' + HEADER_ID + ' .sapUiForm');
            if (!form) return;
            form.querySelectorAll('.sapMLabel').forEach((lab) => {
                const row = lab.closest('.sapUiFormElement') || lab.closest('.sapUiRespGridRow');
                if (!row) return;
                const val = row.querySelector('.sapMText, .sapMTextMaxLine, .sapMObjectNumberText');
                if (!val) return;

                const lr = lab.getBoundingClientRect();
                const vr = val.getBoundingClientRect();
                if (!lr.height || !vr.height) return;

                const rozdiel = (vr.top + vr.height / 2) - (lr.top + lr.height / 2);
                const teraz = parseFloat(lab.style.top) || 0;
                const nove = Math.round(teraz + rozdiel);
                if (Math.abs(nove - teraz) > 1) {
                    lab.style.position = 'relative';
                    lab.style.top = nove + 'px';
                }
            });
        }

        function posun(el, vlastnost, rozdiel) {
            if (!el || !isFinite(rozdiel)) return;
            const teraz = parseFloat(el.style[vlastnost]) || 0;
            const nove = Math.max(0, Math.round(teraz + rozdiel));
            if (Math.abs(nove - teraz) > 1) el.style[vlastnost] = nove + 'px';
        }

        /*
         * Riadok stavovych tlacidiel patri nad box so zakazkou, materialom
         * a vyrobnou zakazkou - operator ich ma mat hned pod nazvom pracoviska.
         * V appke su pod nim, takze ich posunieme na zaciatok Order_FlexBox.
         */
        function tlacidlaHore() {
            const stav = document.getElementById(STATUS_ID);
            const order = document.getElementById(ORDER_ID);
            if (!stav || !order || stav.parentElement !== order) return;
            if (order.firstElementChild === stav) return;
            order.insertBefore(stav, order.firstElementChild);
        }

        function apply() {
            const header = document.getElementById(HEADER_ID);
            if (!header) return;
            injectStyles();
            tlacidlaHore();

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

            zarovnajVlavo();
            zarovnajRiadky();
            oznacAktivitu();
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* ---------- 3.15 Priestorove kolacove grafy (casy SAP) ---------- */

    /*
     * Tri kolace pod popisom operacie (SAP Setup / Machine / Labor Time) su ploche.
     * Modul ich nechava presne take, ake su - nekresli ich odznova, nemeni data -
     * len ich vizualne nakloni (pohlad zboku) a prida tien, takze posobia
     * priestorovo. Pri prechode mysou sa naklon zmensi, akoby sa graf otocil k tebe.
     *
     * Kolace sa hladaju cez texty casov, ktore maju stabilne ID
     * (WorkcenterDetail--SetupTime_Text / MachineTime_Text / LaborTime_Text) -
     * kolac je v tom istom boxe. Vdaka tomu sa netrafi do casoveho grafu vlavo.
     */
    function modDonut3D() {
        const KOTVY = ['WorkcenterDetail--SetupTime_Text',
                       'WorkcenterDetail--MachineTime_Text',
                       'WorkcenterDetail--LaborTime_Text'];
        const MIMO = ['ResourceDetails', 'DialogChart'];
        const STYLE_ID = '__pda_donut3d_styles__';
        const TRIEDA = 'pda-3d';

        /*
         * Vyska (bocna stena) kolaca - lacno.
         *
         * Stena je znovu z `drop-shadow`, ale uz len z TROCH po 4 px namiesto
         * dvadsiatich po 1 px. Kazdy tien sa pocita na vysledok predchadzajuceho,
         * takze 20 tienov = 20 prekresleni obrazka (to mrazilo aplikaciu),
         * kym 3 x 4 px daju rovnako vysoku (12 px) a rovnako plnu stenu za tri.
         * Stena je plna preto, ze posun 4 px je oproti hrubke prstenca maly,
         * takze sa kopie prekryvaju.
         *
         * Druha - dolezitejsia - uspora: `filter` je v pokoji aj pod mysou
         * ROVNAKY a meni sa len `transform`. Prehliadac si tak filtrovany obrazok
         * odlozi a pri naklone uz len posuva hotovu tabulku (robi to graficka
         * karta). Predtym sa pri prechode mysou menil aj filter, takze sa cely
         * obrazok prefiltrovaval znovu v kazdom snimku.
         */
        const STENA = 3;      // kolko tienov (pozor: kazdy dalsi stoji prekreslenie)
        const KROK = 4;       // px na jeden tien -> vyska steny je STENA * KROK

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const stena = new Array(STENA).fill('drop-shadow(0 ' + KROK + 'px 0 rgba(12,28,55,.40))').join(' ');
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/*
 * Stred otacania a zvacsovania je na HORNOM okraji (50% 0).
 * Predtym bol v strede, takze zvacseny kolac rastol aj nahor a pri pretoceni
 * do 2D (kde je najvyssi) siahal na cisla casov nad sebou - posuvanie nadol
 * to riesilo len cast. Takto rastie vylucne nadol a horny okraj sa nehne,
 * takze medzera nad kolacom ostava rovnaka pri kazdom zvacseni.
 */
.${TRIEDA} { transform:perspective(720px) rotateX(38deg) scale(var(--pda3d,1)) !important;
  filter:${stena} drop-shadow(0 12px 9px rgba(16,36,63,.28)) !important;
  transform-origin:50% 0% !important; margin-top:22px !important;
  transition:transform .24s ease !important;
  will-change:transform !important; }
/* pri prechode mysou sa kolac pretoci do skutocneho tvaru - filter ostava rovnaky */
.${TRIEDA}:hover { transform:perspective(720px) rotateX(0deg)
    scale(calc(var(--pda3d,1) * 1.06)) !important; }
/* aby naklonený kolac ani jeho stena neboli orezane okrajom boxu */
.pda-3d-box { overflow:visible !important; }
`;
            document.head.appendChild(st);
        }

        function box(el) {
            return el.closest('.sapMVBox') || el.closest('.sapMFlexBox') || el.parentElement;
        }

        /*
         * Na sirokej obrazovke su kolace zbytocne male - appka ich kresli stale
         * rovnako velke. Zvacsujeme ich zvacsenim (transform), takze sa nemeni
         * ziadny rozmer v rozlozeni a nic sa nikam neposunie; miesta okolo nich
         * je dost.
         */
        function zvacsenie() {
            const w = W.innerWidth;
            if (w >= 2200) return 1.9;
            if (w >= 1900) return 1.65;
            if (w >= 1650) return 1.4;
            if (w >= 1400) return 1.2;
            return 1;
        }

        function apply() {
            injectStyles();
            KOTVY.forEach((id) => {
                const text = document.getElementById(id);
                if (!text) return;
                const b = box(text);
                if (!b) return;
                if (!b.classList.contains('pda-3d-box')) b.classList.add('pda-3d-box');

                b.querySelectorAll('canvas, svg').forEach((g) => {
                    if (MIMO.indexOf(g.id) !== -1) return;
                    if (g.classList.contains(TRIEDA)) return;
                    // ikonky a drobnosti nechame tak, kolac je velky
                    const r = g.getBoundingClientRect();
                    if (r.width < 60 || r.height < 60) return;
                    g.classList.add(TRIEDA);
                });

                const z = String(zvacsenie());
                b.querySelectorAll('.' + TRIEDA).forEach((g) => {
                    if (g.style.getPropertyValue('--pda3d') !== z) g.style.setProperty('--pda3d', z);
                });
            });
        }

        DomWatch.add(apply);
        W.addEventListener('resize', apply);
        onReady(apply);
    }

    /* ---------- 3.16 Lavy panel na celu vysku obrazovky ---------- */

    /*
     * Povodne rozlozenie: cez celu sirku ide panel "Osobny stav", pod nim panel
     * pracoviska a az v nom vlavo zoznam zakaziek. Zoznam tak zacina az v polovici
     * obrazovky a hore vlavo je velka prazdna plocha.
     *
     * Modul to preskladá bez zasahu do appky - iba polohovanim:
     *   - lavy stlpec (zoznam zakaziek + graf) sa pripne nalavo od horneho
     *     baneru (HF logo, meno) az po spodok obrazovky
     *   - panel "Osobny stav" aj panel pracoviska dostanu zlava odsadenie,
     *     takze uz nezacinaju pri lavom okraji, ale az za tym stlpcom
     *
     * Nic sa nepresuva v DOM, len sa nastavuje poloha - ked sa modul vypne,
     * appka je presne taka, aka bola.
     */
    function modFullHeightLayout() {
        const LEFT_ID = 'WorkcenterDetail--LeftColumn_FlexBox';
        const STYLE_ID = '__pda_layout_styles__';
        const SIRKA = 430;      // sirka laveho stlpca (rovnaka ako v module zoznamu)
        const MEDZERA = 14;     // medzera medzi stlpcom a obsahom vpravo
        const SPODOK = 10;      // medzera pod stlpcom

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
.pda-layout-left { position:fixed !important; z-index:3 !important;
  width:${SIRKA}px !important; max-width:${SIRKA}px !important; flex:0 0 ${SIRKA}px !important;
  display:flex !important; flex-direction:column !important; box-sizing:border-box !important; }
.pda-layout-odsad { margin-left:${SIRKA + MEDZERA}px !important; }
/* box so zoznamom si vezme vsetko volne miesto, graf ostane taky, aky je */
.pda-layout-left > .pda-left-box:first-of-type { flex:1 1 auto !important; min-height:0 !important; }
.pda-layout-left > .pda-left-box:last-of-type { flex:0 0 auto !important; }
`;
            document.head.appendChild(st);
        }

        // panel, v ktorom su tlacidla osobneho stavu (Stretnutia / Prestavka / cakanie)
        function osobnyPanel(left) {
            const btn = Array.from(document.querySelectorAll('.statusBtn'))
                .find((b) => !left.contains(b) && b.closest('.sapMPanel'));
            return btn ? btn.closest('.sapMPanel') : null;
        }

        function apply() {
            const left = document.getElementById(LEFT_ID);
            if (!left || !left.offsetParent && left.style.position !== 'fixed') {
                // detail pracoviska nie je otvoreny - nic neriesime
                if (!left) return;
            }
            injectStyles();

            const osobny = osobnyPanel(left);
            const detail = left.closest('.sapMPanel');
            if (!osobny || !osobny.parentElement) return;

            // odsadenie oboch panelov vpravo od stlpca
            [osobny, detail].forEach((p) => {
                if (p && !p.classList.contains('pda-layout-odsad')) p.classList.add('pda-layout-odsad');
            });

            /*
             * Horny okraj stlpca = horny okraj panela "Osobny stav", teda presne
             * pod banerom s logom HF. Lavy okraj sa berie z rodica panela, nie
             * zo samotneho panela - ten sme prave posunuli doprava, takze by
             * sme merali uz posunutu hodnotu.
             */
            const rodic = osobny.parentElement.getBoundingClientRect();
            const hore = Math.round(osobny.getBoundingClientRect().top);
            const vlavo = Math.round(rodic.left + 8);
            if (!isFinite(hore) || hore < 0) return;

            if (!left.classList.contains('pda-layout-left')) left.classList.add('pda-layout-left');
            if (left.style.top !== hore + 'px') left.style.top = hore + 'px';
            if (left.style.left !== vlavo + 'px') left.style.left = vlavo + 'px';
            if (left.style.bottom !== SPODOK + 'px') left.style.bottom = SPODOK + 'px';
        }

        DomWatch.add(apply);
        W.addEventListener('resize', apply);
        onReady(apply);
    }

    /* ---------- 3.17 Menu HF Slovakia (pravy bocny panel) ---------- */

    /*
     * Zvisly panel pri pravom okraji s pripravovanymi funkciami HF Slovakia.
     * Vpravo od obsahu appky (za blokom Components / BOM) je volne miesto,
     * takze panel nic neprekryva a hlavne: je to NAS vlastny prvok pripnuty
     * k oknu - z rozlozenia appky sa nic nevybera a nic sa nepresuva.
     * (Pokus pripnut lavy stlpec appky v 1.22.0 rozlozenie rozhodil, preto
     * sa tu appky nedotykame vobec.)
     *
     * Tlacidla zatial nic nerobia - po kliknuti sa ukaze okno "vo vyvoji".
     * Funkcie sa budu doplnat postupne, kazda do svojej vetvy v `spusti`.
     */
    function modHfMenu() {
        const PANEL_ID = '__pda_hf_menu__';
        const OVERLAY_ID = '__pda_hf_overlay__';
        const STYLE_ID = '__pda_hf_menu_styles__';
        const SIRKA = 250;
        const OKRAJ = 10;
        const TAB_ID = '__pda_hf_tab__';
        const KEY_OTVORENY = 'pda_hfmenu_open_v1';
        const PRAH = 1600;      // od akej sirky okna je panel predvolene otvoreny

        /*
         * Na uzkej obrazovke panel prekryval stavove tlacidla, preto sa da
         * schovat: vpravo ostane uzky pasik a klik na neho panel vysunie.
         * Predvolene je otvoreny len na sirokej obrazovke; kedze si volbu
         * pamatame, na dielni sa nastavi raz a drzi aj po obnoveni stranky.
         */
        let otvoreny = null;

        function nacitajStav() {
            try {
                const v = GM_getValue(KEY_OTVORENY, null);
                if (v === '1') return true;
                if (v === '0') return false;
            } catch (e) { /* ignore */ }
            return null;
        }

        function ulozStav(v) {
            try { GM_setValue(KEY_OTVORENY, v ? '1' : '0'); } catch (e) { /* ignore */ }
        }

        const LOGO_HF = 'data:image/jpeg;base64,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';
        const LOGO_FLEXUS = 'data:image/jpeg;base64,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';
        const FOTO_ELA = 'data:image/jpeg;base64,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';

        const POLOZKY = [
            { ikona: '📐', nazov: 'Výkresy', popis: 'nájsť podľa čísla', akcia: 'vykresy' },
            { ikona: '🌀', nazov: 'CHIPS', popis: 'vývoz špon' },
            { ikona: '👷', nazov: 'Privolanie majstra', popis: 'majster na pracovisko' },
            { ikona: '🚚', nazov: 'Odviezť materiál', popis: 'odvoz hotových dielov' },
            { ikona: '🔧', nazov: 'Privolanie TOOLSHOP', popis: 'nástrojáreň' },
            { foto: FOTO_ELA, nazov: 'Ela', popis: 'elektronická asistentka (AI)' },
            { obrazok: LOGO_FLEXUS, nazov: 'Flexus', lenObrazok: true },
        ];

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* Nizke z-index je zamerne: panel ma byt nad obsahom stranky, ale POD
   oknami appky (detail grafu, tabulka stavov). S vysokym cislom prekryval
   otvorene okno a nestmavil sa spolu so zvyskom stranky. */
#${PANEL_ID} { position:fixed; right:${OKRAJ}px; width:${SIRKA}px; z-index:5;
  display:flex; flex-direction:column; gap:8px; box-sizing:border-box;
  font:13px/1.35 -apple-system,"Segoe UI",Roboto,sans-serif; color:#13315c;
  padding:10px; background:#fff; border:1px solid #dfe4ec; border-radius:16px;
  box-shadow:0 4px 18px rgba(16,36,63,.12); overflow-y:auto; justify-content:flex-start; }
#${PANEL_ID} .hf-logo { width:100%; border-radius:10px; display:block; }
#${PANEL_ID} .hf-nadpis { font-size:10.5px; font-weight:800; letter-spacing:.14em;
  text-transform:uppercase; color:#8e9bb0; text-align:center; margin:2px 0 4px; }
#${PANEL_ID} .hf-btn { display:flex; align-items:center; gap:10px; width:100%; box-sizing:border-box;
  padding:10px 12px; border:2px solid #c9d7ea; border-radius:12px;
  background:linear-gradient(180deg,#f7faff 0%,#eaf1fb 100%); cursor:pointer; text-align:left;
  font:inherit; color:#13315c; box-shadow:0 2px 6px rgba(16,36,63,.10);
  transition:transform .13s ease, box-shadow .13s ease, border-color .13s; }
#${PANEL_ID} .hf-btn:hover { transform:translateY(-3px); border-color:#2563eb;
  box-shadow:0 10px 20px rgba(16,36,63,.22); }
#${PANEL_ID} .hf-btn:active { transform:translateY(-1px); box-shadow:0 3px 8px rgba(16,36,63,.18); }
#${PANEL_ID} .hf-btn .ik { font-size:22px; line-height:1; flex:0 0 auto; width:26px; text-align:center; }
#${PANEL_ID} .hf-btn .ik img { width:100%; display:block; border-radius:3px; }
/* fotka Ely je okruhla a o nieco vacsia ako ikonky */
#${PANEL_ID} .hf-btn .ik.foto { width:34px; }
#${PANEL_ID} .hf-btn .ik.foto img { border-radius:50%; border:2px solid #c9d7ea; }
#${PANEL_ID} .hf-btn .tx { flex:1 1 auto; min-width:0; }
#${PANEL_ID} .hf-btn .n { display:block; font-size:13px; font-weight:700; }
#${PANEL_ID} .hf-btn .p { display:block; font-size:11px; color:#6b7c95; margin-top:1px;
  overflow:hidden; text-overflow:ellipsis; white-space:nowrap; }
/* tlacidlo, na ktorom je len logo - bez textu */
#${PANEL_ID} .hf-btn.hf-obr { justify-content:center; padding:10px 14px; }
#${PANEL_ID} .hf-btn.hf-obr img { width:100%; max-width:175px; display:block; border-radius:4px; }

/* Zatvaranie a otvaranie je posunom za pravy okraj, nie zmiznutim - je tak
   vidiet, kam sa panel podel. Viditelnost sa prepne az na konci prechodu,
   aby schovany panel nechytal kliknutia. */
#${PANEL_ID} { transition:transform .28s ease, opacity .28s ease, visibility 0s linear .28s; }
#${PANEL_ID}.skryty { transform:translateX(calc(100% + ${OKRAJ + 6}px)); opacity:0;
  visibility:hidden; pointer-events:none; }
/* uzky pasik na pravom okraji - klikom sa panel vysunie alebo schova */
#${TAB_ID} { position:fixed; z-index:5; width:26px; padding:14px 0; cursor:pointer;
  transition:right .28s ease;
  display:flex; flex-direction:column; align-items:center; gap:8px;
  background:#13315c; color:#fff; border:0; border-radius:10px 0 0 10px;
  box-shadow:-2px 2px 10px rgba(16,36,63,.25); font:inherit; }
#${TAB_ID}:hover { background:#1c478a; }
#${TAB_ID} .sip { font-size:13px; line-height:1; }
#${TAB_ID} .txt { writing-mode:vertical-rl; text-orientation:mixed; font-size:10.5px;
  font-weight:800; letter-spacing:.10em; white-space:nowrap; }

#${OVERLAY_ID} { position:fixed; inset:0; background:rgba(10,20,40,.5); z-index:100002;
  display:flex; align-items:center; justify-content:center; padding:24px; box-sizing:border-box; }
#${OVERLAY_ID} .karta { background:#fff; border-radius:16px; width:min(460px,92vw); overflow:hidden;
  box-shadow:0 24px 70px rgba(16,36,63,.4); font-family:-apple-system,"Segoe UI",Roboto,sans-serif; }
#${OVERLAY_ID} .hl { display:flex; align-items:center; gap:12px; padding:14px 18px;
  background:#13315c; color:#fff; }
#${OVERLAY_ID} .hl .n { font-size:15px; font-weight:800; letter-spacing:.06em; text-transform:uppercase; }
#${OVERLAY_ID} .hl .x { margin-left:auto; background:none; border:0; color:#fff; font-size:28px;
  line-height:1; cursor:pointer; padding:0 4px; }
#${OVERLAY_ID} .telo { padding:22px 20px 24px; font-size:15px; line-height:1.6; color:#17202e; text-align:center; }
#${OVERLAY_ID} .telo .velke { font-size:38px; display:block; margin-bottom:10px; }
#${OVERLAY_ID} .telo .co { font-weight:700; color:#13315c; }
#${OVERLAY_ID} .telo.hladat { text-align:left; }
#${OVERLAY_ID} .riadok { display:flex; gap:8px; margin-top:12px; }
#${OVERLAY_ID} input.pole { flex:1 1 auto; min-width:0; padding:10px 12px; font-size:16px;
  border:2px solid #c9d7ea; border-radius:10px; color:#13315c; box-sizing:border-box; }
#${OVERLAY_ID} input.pole:focus { outline:none; border-color:#2563eb; box-shadow:0 0 0 3px rgba(37,99,235,.18); }
#${OVERLAY_ID} button.hladaj { flex:0 0 auto; padding:10px 20px; font-size:15px; font-weight:700;
  border:0; border-radius:10px; background:#13315c; color:#fff; cursor:pointer;
  box-shadow:0 2px 8px rgba(16,36,63,.22); }
#${OVERLAY_ID} button.hladaj:hover { background:#1c478a; }
#${OVERLAY_ID} .napoveda { font-size:12.5px; color:#6b7c95; margin-top:10px; }
`;
            document.head.appendChild(st);
        }

        function okno(nazov) {
            const stary = document.getElementById(OVERLAY_ID);
            if (stary) stary.remove();

            const overlay = document.createElement('div');
            overlay.id = OVERLAY_ID;

            const karta = document.createElement('div');
            karta.className = 'karta';

            const hl = document.createElement('div');
            hl.className = 'hl';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = 'HF Slovakia';
            const x = document.createElement('button'); x.className = 'x'; x.type = 'button';
            x.textContent = '×'; x.setAttribute('aria-label', 'Zavrieť');
            hl.appendChild(n); hl.appendChild(x);

            const telo = document.createElement('div');
            telo.className = 'telo';
            const ik = document.createElement('span'); ik.className = 'velke'; ik.textContent = '🚧';
            const co = document.createElement('div'); co.className = 'co'; co.textContent = nazov;
            const txt = document.createElement('div');
            txt.textContent = 'Táto funkcia je vo vývoji. Zatiaľ nie je funkčná — pripravujeme ju.';
            telo.appendChild(ik); telo.appendChild(co); telo.appendChild(txt);

            karta.appendChild(hl); karta.appendChild(telo);
            overlay.appendChild(karta);

            const zavri = () => { overlay.remove(); document.removeEventListener('keydown', naEsc); };
            const naEsc = (e) => { if (e.key === 'Escape') zavri(); };
            x.addEventListener('click', zavri);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) zavri(); });
            document.addEventListener('keydown', naEsc);
            document.body.appendChild(overlay);
        }

        /*
         * Zadanie cisla vykresu. Samotne hladanie a zoznam najdenych suborov
         * uz robi okno z modulu vykresu (`shared.pdmOpenDialog`) - je to to iste
         * okno, ake sa otvori kliknutim na okienko VYKRES pri zakazke, takze
         * sa spravanie nikde nerozchadza.
         */
        function oknoVykresy() {
            if (typeof shared.pdmOpenDialog !== 'function') {
                okno('Výkresy');
                return;
            }

            const stary = document.getElementById(OVERLAY_ID);
            if (stary) stary.remove();

            const overlay = document.createElement('div');
            overlay.id = OVERLAY_ID;

            const karta = document.createElement('div');
            karta.className = 'karta';

            const hl = document.createElement('div');
            hl.className = 'hl';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = 'Výkresy';
            const x = document.createElement('button'); x.className = 'x'; x.type = 'button';
            x.textContent = '×'; x.setAttribute('aria-label', 'Zavrieť');
            hl.appendChild(n); hl.appendChild(x);

            const telo = document.createElement('div');
            telo.className = 'telo hladat';
            const popis = document.createElement('div');
            popis.textContent = 'Zadaj číslo výkresu alebo číslo materiálu:';

            const riadok = document.createElement('div');
            riadok.className = 'riadok';
            const pole = document.createElement('input');
            pole.className = 'pole';
            pole.type = 'text';
            pole.placeholder = 'napr. 1-60.2-06.66-009 alebo 25217930';
            const hladaj = document.createElement('button');
            hladaj.className = 'hladaj';
            hladaj.type = 'button';
            hladaj.textContent = 'Hľadať';
            riadok.appendChild(pole); riadok.appendChild(hladaj);

            const napoveda = document.createElement('div');
            napoveda.className = 'napoveda';
            napoveda.textContent = 'Zobrazí sa zoznam dostupných výkresov — kliknutím na riadok sa výkres otvorí.';

            telo.appendChild(popis); telo.appendChild(riadok); telo.appendChild(napoveda);
            karta.appendChild(hl); karta.appendChild(telo);
            overlay.appendChild(karta);

            const zavri = () => { overlay.remove(); document.removeEventListener('keydown', naEsc); };
            const naEsc = (e) => { if (e.key === 'Escape') zavri(); };
            const odosli = () => {
                const cislo = pole.value.trim();
                if (!cislo) { pole.focus(); return; }
                zavri();
                shared.pdmOpenDialog(cislo, { titul: cislo });
            };

            x.addEventListener('click', zavri);
            overlay.addEventListener('click', (e) => { if (e.target === overlay) zavri(); });
            hladaj.addEventListener('click', odosli);
            pole.addEventListener('keydown', (e) => { if (e.key === 'Enter') odosli(); });
            document.addEventListener('keydown', naEsc);

            document.body.appendChild(overlay);
            setTimeout(() => pole.focus(), 30);
        }

        // sem budu pribudat skutocne funkcie - kazda si najde svoju vetvu podla `akcia`
        function spusti(polozka) {
            if (polozka.akcia === 'vykresy') { oknoVykresy(); return; }
            okno(polozka.nazov);
        }

        function tlacidlo(p) {
            const b = document.createElement('button');
            b.type = 'button';
            b.className = 'hf-btn';

            // Flexus ma na tlacidle len svoje logo, ziadny text
            if (p.lenObrazok && p.obrazok) {
                b.classList.add('hf-obr');
                const img = document.createElement('img');
                img.src = p.obrazok;
                img.alt = p.nazov;
                b.appendChild(img);
                b.title = p.nazov;
                b.addEventListener('click', () => spusti(p));
                return b;
            }

            const ik = document.createElement('span');
            ik.className = 'ik';
            if (p.foto) {
                ik.classList.add('foto');
                const img = document.createElement('img');
                img.src = p.foto;
                img.alt = p.nazov;
                ik.appendChild(img);
            } else if (p.obrazok) {
                const img = document.createElement('img');
                img.src = p.obrazok;
                img.alt = p.nazov;
                ik.appendChild(img);
            } else {
                ik.textContent = p.ikona;
            }

            const tx = document.createElement('span');
            tx.className = 'tx';
            const n = document.createElement('span'); n.className = 'n'; n.textContent = p.nazov;
            tx.appendChild(n);
            if (p.popis) {
                const s = document.createElement('span'); s.className = 'p'; s.textContent = p.popis;
                tx.appendChild(s);
            }

            b.appendChild(ik); b.appendChild(tx);
            b.addEventListener('click', () => spusti(p));
            return b;
        }

        function postav() {
            const panel = document.createElement('div');
            panel.id = PANEL_ID;

            const logo = document.createElement('img');
            logo.className = 'hf-logo';
            logo.src = LOGO_HF;
            logo.alt = 'PDA App Extension — developed by HF Slovakia';
            panel.appendChild(logo);

            const nad = document.createElement('div');
            nad.className = 'hf-nadpis';
            nad.textContent = 'Funkcie HF Slovakia';
            panel.appendChild(nad);

            POLOZKY.forEach((p) => panel.appendChild(tlacidlo(p)));
            document.body.appendChild(panel);
            return panel;
        }

        /*
         * Panel patri LEN na obrazovku otvoreneho pracoviska - na uvodnej
         * obrazovke (prehlad pracovisk) nema co robit. Pozname ju podla toho,
         * ci je vidiet panel s pracovnym zoznamom.
         */
        function panelPracoviska() {
            const left = document.getElementById('WorkcenterDetail--LeftColumn_FlexBox');
            const p = left && left.closest('.sapMPanel');
            if (!p) return null;
            const r = p.getBoundingClientRect();
            return r.height > 0 && r.width > 0 ? p : null;
        }

        function pasik() {
            let t = document.getElementById(TAB_ID);
            if (t) return t;
            t = document.createElement('button');
            t.id = TAB_ID;
            t.type = 'button';
            const sip = document.createElement('span'); sip.className = 'sip';
            const txt = document.createElement('span'); txt.className = 'txt'; txt.textContent = 'HF SLOVAKIA';
            t.appendChild(sip); t.appendChild(txt);
            t.addEventListener('click', () => {
                otvoreny = !otvoreny;
                ulozStav(otvoreny);
                apply();
            });
            document.body.appendChild(t);
            return t;
        }

        function apply() {
            injectStyles();
            const detail = panelPracoviska();
            const panel = document.getElementById(PANEL_ID);
            const t = document.getElementById(TAB_ID);

            // mimo detailu pracoviska nie je vidiet ani panel, ani pasik
            if (!detail) {
                if (panel) panel.style.display = 'none';
                if (t) t.style.display = 'none';
                return;
            }

            if (otvoreny === null) {
                const ulozene = nacitajStav();
                otvoreny = ulozene === null ? W.innerWidth >= PRAH : ulozene;
            }

            const p = panel || postav();
            p.style.display = '';
            p.classList.toggle('skryty', !otvoreny);

            const tab = pasik();
            tab.style.display = '';
            tab.title = otvoreny ? 'Skryť panel HF Slovakia' : 'Zobraziť panel HF Slovakia';
            tab.querySelector('.sip').textContent = otvoreny ? '▶' : '◀';
            tab.style.right = otvoreny ? (SIRKA + 2 * OKRAJ) + 'px' : '0px';

            /*
             * Zaciatok je zarovnany s panelom pracoviska (teda pod tlacidlami
             * Stretnutia / Prestavka) a dole siaha az k spodku okna, takze
             * vyplni celu volnu plochu vpravo.
             */
            const hore = Math.round(detail.getBoundingClientRect().top);
            if (hore > 0 && Math.abs((parseFloat(p.style.top) || 0) - hore) > 4) {
                p.style.top = hore + 'px';
                p.style.bottom = OKRAJ + 'px';
            }
            if (hore > 0) tab.style.top = (hore + 12) + 'px';
        }

        DomWatch.add(apply);
        W.addEventListener('resize', apply);
        onReady(apply);
    }

    /* ---------- 3.18 NOVY DIZAJN (vetva ver.2) ---------- */

    /*
     * Prezlecenie PDA do dizajnu HF Slovakia podla navrhu z 2026-09-16.
     *
     * Zasada, na ktorej je to postavene: NIC z appky sa nemaze ani nevybera
     * z toku stranky. Vsetko su len styly a par vlastnych prvkov vlozenych
     * medzi povodne. Tlacidla, prepinace a zoznamy ostavaju tie ISTE prvky
     * appky, takze vsetky volania a funkcie fungujú presne ako predtym -
     * menia sa iba farby, tvary a popisy.
     * (Pokus pripnut lavy stlpec cez position:fixed vo vetve 1.22 rozlozenie
     * rozhodil - preto sa tu nic nepolohuje natvrdo.)
     *
     * Cast 1: pozadie a karty, horny pruh s logom, stavove tlacidla ako karty
     * s ikonou a podnadpisom, nadpisy sekcii, patka.
     */
    function modNewDesign() {
        const STYLE_ID = '__pda_nd_styles__';
        const BODY_CLASS = 'pda-nd';
        const HEADER_ID = '__pda_nd_header__';
        const FOOTER_ID = '__pda_nd_footer__';
        const LOGO_HF_MALE = 'data:image/jpeg;base64,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';
        const POZADIE = 'data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAcFBgYGBQcGBgYICAcJCxIMCwoKCxcQEQ0SGxccHBoXGhkdISokHR8oIBkaJTIlKCwtLzAvHSM0ODQuNyouLy7/2wBDAQgICAsKCxYMDBYuHhoeLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi4uLi7/wAARCAOtBogDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD6JAApaKKskKKKKACiiigAooooAKKKKAEopaSgAooooAKKKKAEooooAKDRRTAaaKWkoAKKKSgANH40UZoEBoopKAFopKWmAZopKM0gCikopgFFJS0ALRSUUAFFJSigQUUUUDCkNFGaBBRSZopgLSUUUALRSUZoAWiiqGsXf2a1IU4lk+Vfb1NOKbdkJuyuY2s3ZubnYpzFFwPc9zVAUgFOFd6SSsjlbu7sUU6milzSAQ0lLSGgBppppxpY0aSRY0GWY4A96YjR0O1DzG7kGY4eQPVuwrXYliWY8nk0gjW2hjtYzkRj5j/ebuaTNcspczudEVZWFopM0VIwozSZpKAAmkNFBpgJSUUUCDNOijaaRY16k02tKxi8qEzMPnfhfYUpOyGlcnYKoWNPuoMCmmikNZFiGmsQoJPQUtNVPOlCfwL8z/0FMCW2Qqhkb77/AKCpqCcmioKsFFJRQAtFJRQAUtJRQAtJRS0AFFFFABRRRQAtFJRQMWikooAKWikoAKKKKACiiigQtFFFAwpaSigBaKSigBaKTNFAhRRRRQMWikopALRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFACUtFFACUtFFABRRRQAUUUUAFFFFABQaKKACiiigAooooAKKKKACiiigAooooAKKDRigAooooAKKKKACiiigAooFFABRRRQAUUUUAFFFFABRRRQACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBk0gjQt37fWsw5JJJyT1NS3MvmScfdHAqKtYqyM5O4lLRS1QgpKKKACiiigBDRRRQAUUUUAJS0UUCFpKKKACiiigAopKKYGtSUtJWBqLRRRQAUUUlAC0UUUAFFJRQAtJRSFgBzwKAFoqBrmBfvP+hpovbX/AJ6/pTswuizSVEtzA33ZkP41KOeRQAtJS0lIAoopKYBSGlpKAEooooEFFFFACUUUUAFFJRTAWkzRSUALSUUUCCikFLQAtJRSUALRmkopgLRmkooAWkNJRQAtJRSUALmlpKKAFopM0UABIUEk4A5JrkdQujd3TSfwDhB7Vr69d+XELVD87jLey/8A165+uqhCy5mYVZXdhadTaWtjIWlpKXNIYUhooNAhprY0S3MatfOOmViHv3P4Vm20D3M6Qx/ec4z6e9dG+xQkUXEUY2r/AI1nVlpY0hHW42iikrE1A0UlFIBc0lJmigApKKSmAUlLSUCJ7WEzTBP4erH2rTkYM3HCjgVHbxiCDB/1j8n2FOrJu7NErISkNLTTSAa7bVJNWYY/Kj2n75+Zj71BAvmS7z9yPp7tVmh9hoKKKKkYlFFBoAKKKKAFpKKKACilpKACiiigApaSigBaKKKACiiigApaSigYUUUUCCiigUDFopCaKAFopKKAFooooAKKKKAFopKKAFopKKAFopKWgApaSlpAFFJS0AFFFFABRRRQAUUUUAFFFLQAlFFFABRRRQAUUUUAFFFBoAKKKKACiiigAooooAKKKKACiiigAooooAKKWkoAWkoooAKKOKKACiiigAooozQAUUUUAFFH4UUAFFFFABRRQaACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKPwoAKr3cm1dg6t19hU0jhELN0FZjsXYs3U1UVcmTG0tFFakBS0lLQAUlLRQAlIaWigBtLRRQIKKKKACkpaKAEopaSgBaKSloAKKKKYGpS0lFYGotJRRQAtFJS0AJS0VBPOI/lXlv5UASu6oMscCqsl4Bwi/iarOzMcscmo2q1Elska5mJ/1hH0qM3Mw6SNUbUwmqsTcsC+uB1KsP9paXz7eTie3A/2kqpnNIc07ILllrGOYE2k6sf7jcGq5NzatjLxn0qIsVOQeR3FWY9T+XyrtBNH6n7wp6+otCWDVWU4uEyP7y9fyrThninTfE4Yd8dRWBd2wMf2i0fzYe4HVaz4riSGQSROVYdxS5FLYfO1udlRWZp2ppckRTAJL29G/+vWnWbTWjLTuJRRSUgFpKKKAEooopgJRRSUCCiiigApDRSZpgLRmkpKAHUZptGaAHUU2igBTRSUUALRTc0ZoAWjNJRQAuaKSigBaKTNFAC0yeVYYnlc/KoyadWFrl3vcWqH5U5f3PpVwjzOxMpWVzLuJXnmeV/vMc/SmUtBrtOUKKKKQxaKKKQBmk70VYsbY3VysXRPvO3oo60N21C1zU0uDyLU3Df6yYbU9l7n8asUsj73yBhQMKPQdqbXO9Xdm600FzRSUlIYtIaKSgBaTNFJQIWkpKWgAzVqwhEsu9h8icn3qqASQqjJPArXVBBCsI69WPqamT0sVFCu25i1JmikrMoM0yQnhVGWbgUpNPtVyTOeg+VP6mnsBMqCNFjXnb1PqaWikqCgoopKAFozSUUwFopKKQC0UlFAC0UUUAFAoooAKKKKAClpO1FAC0UUUAFFFFABRRRQMKKKDQIKKKKBh0ooooAWikooAWlpKKACiiigBKWkpaACiiigApaKSkAtFFFAC0UlFAC0UUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFHFFHFABRRRQAUUUUAFGaKKACiiigAooooAKKKKACiijNABRQKOaAFpKKWgBKKMUUAFFFFABRRRQAUUUUAFFFFABiiiigAooooAKKKKADNFFFABRRRQAUUGigAooooAKKKKADFLSVBcy+WmB95untTSuBBdyb32A/Kv6mq9FLWyVjK9xKKKKAClpKKAFpKKKACiiigApDS0lABRRRQAoopKKAFpKWkoAKKKKACijNFMRqUUUVgahS0lFAC0lFMmkEaFu/agCO5m2Dav3j+lUiaCxJJY8mm5q0iWwJoVWkbai5NLFE0r4HAHU1oxosa7VGBTbsJK5WSyHWRvwFTrbwL0jH481LRUtsqw3ao6KB+FGF7qp+opaDSGV5bW3lHzQqD6jis260nIJgk/4C/wDjWzijHrVKTRLSZysZu9PnztKHup6MKtXFnHdQG9tFwR/rI+6muheKKZCkqKynsax54JtKm+1WxMkH8a98e/8AjVqd/Unlt6GJjFbmlaj5xFtO370fdY/xe31rO1VI123Vvzbzcj/ZPpWUJGDBlOGByCO1acvOiL8rO7oqlpd4Ly2DnAkXhx7+tXK52rOxqncKKKSgYUUlFABRSUUCCikoNMAopKSgBaSikzQAtGaKKACiiigAopM0ZoAWkoopgFGaSigBTRSUZoAWgUlFAiG9uBbW7SfxdFHqa5VssxYkkk5Jq7qlz9ouCqn92nA9z3NUhXXTjyo55yuwxRQaKskSijFLigApMUtFIBproLKH7LZAMMTT4Zv9lew/rWdpVstxcb5B+5iG9/f0H41qyOZHZ26k1lUd9DSC6iUUmaKgsWikpKAFNJRmkzSAWkopKYBRmkqSGNppVjXqxoAu6dEPmuHHC8L7mrJOSSetObairEn3EGKZmsb31NLW0A03NKaaxwMnpQA1gXZYlPzN39B3NXMAAKvCqMCobZNqGZh8z9PZalpNjQtJRRSGFFJRQAtJRRQAUtJRQAtJRRQAtFJS5oAKKKKAClpKBSAWikpaACikooAdRSUUALRSUUALRSUUALRRRQMKKKKACiiigAoooxQAtFJS0AJRRRQAUtJRQAtFJRQAtFFFAC0UlFAC0tJQKQC0UUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRiiigAooooAKKKKACiiigAoooFABRRRQAUUUUAFFFFABRRRQAUUCigAooNFABR1ozRQAUtJR9KAA0UUUAFFFFACMQqlj0HWsyRzI5Y/hVi8lyfLB4HWqtaxVtSJMKKKKokKKKKACiiikAUUUUwCiiigApKWkoAWkpaSgAooooAKKKKACloooADRRRQI06KBS1iaiUUUUAFRTQiUgliMdhUhOAT6Csc3lwST5hHsBVJNibLxtF/vtSfYl/vtVH7Xcf89WqzYXE0lxsdyy4PBp2aFoXY41jTav5+tPpTimk1JQtJWVq11PFKkcT7AVySOtURfXf/AD8PVKDauS5JHR0VzRv7vI/0huorpMjA+lKUWgTTFNJkCop5GSGR1GSqkgVy7X983P2hhnnAxVRhcUpWOs3D1prHIxXJG/vh/wAvL/pTRqF8Ol1J+dX7Jk850Q0+EQzQDPlSHIX+4faqa+H4P+e8v5Csoanfj/l6f9KDq1+OPtT/AKU1GfRivHqb9lpa2UxljnkORgqQMGr9ca2r6h/z9yfpXQ6LdS3Vgskx3OGKlvXFTOElqyoyWyNGkopKzKFzSUUlMAoopKAClptGaAFpKKSgAopKWmAUUZpKAFopOKKBBmkopKAFzRSUUAOpM0lFAC5opKWgAqlqlz5FvsU4kk4HsO5q47BFLMcKBkmuZu52uJ2lPToo9BWtOF3cicrIgIoxxS0V0mA2ilooAKKSjNIANIM5wBmlrS0eAF2u5BlIvug937fl1pN2VxpXdi9FD9ltktv4z80p/wBr0/Cigkkkk5J60lYGotJmkooADRmkNFABmikozQAtNNGaKAErX0+LyYDOw+eThfYVn2cBuJwn8I5Y+1asr7m44UcCom+hcV1EpDRmkNQUGaaq+dKI/wCEcv8ASkdgqk1YgQxRYb/WP8zf4UbASMcn27U2lpM1JQUUUUAFFFJQAtFFFABRSUUALRSUUAFLSUtAAKWkpaACiikoAWiiigAooooAWiikoAWikpaQBRRRQAUUUUALRRRQAlFFFAxaKKKBBRRRQMKWkpaAEooooAKWkooAWiiigAooooAWgUlFAC0UUUgFopKWgAooooAKKSigBaKKSgBaKKKACiiigAooooAKKKKACiiigAooooAWkoooAKKKKACiiigAooxRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUALSUUUAFFFFABRRRQAtRTyeWmf4j0qQ4AyelZ00hkkLdh0qoq7E3Yj60UUVqZhSUUUAFFFFABRRRSAKKKKACiiigAopKKYBRRRQAUUppKAFopKWgAoopRQAlFFFAjTooorE1ClpKWgBr/cb6GueNdC/wBxvof5VzZNaQJkOzVrTD/pQ/3TVImrel/8fY/3TVNaErc2qaRS0dayLMLWv+PiP/c/rWfmtHWsG6UeiCs01vHYyluNxlh9RXXBeB9BXJA/MPqP511wPA+lTU6FQIrkD7NN/uGuPJGBXX3Z/wBFm/3DXGZ4FOlsKYretMNOOa0tDt4Z5ZTMgcIBgHpWjdlchK5kmomPNdqbGyP/AC6xf980n9nWP/PrF/3zU+1Q+RnDkZrq/DYxpg/66N/Srv8AZ1j/AM+kX/fNTxxxxII4kVEHQKOKmdRSViows7j6SikrIsDRRSUwFpKSjNABRRmjNABTaKKYBRSUUCFpM0lJQA7NFNzRQAtFJRTAKKKSgBaWkpRQAtFJUc8qwxNK3RR09TQkBnazc4AtkPJ5f+grIqSRmkdpH5Zjk0w11xXKrHPJ3dxKKWkpkhQaKDQA00hp1NNADo1aSRY0GWY4A9TXQsqwxpbRnKxjk/3m7mqOkQ+VG14w+Y5SL69zVmspu7NIqyFoopKkoXNJSGigBaSkooADSUUlAC0UVd02ASSmV/uR8/U0m7K4JXdi5bxfZrYKf9ZJy3sPSinOxdixplYmotJRTHYgAKMsxwo96AHwoJJtzfcj5+pqwTk5pFURRrEDnHJPqaKlu4xaKSloASlpKM0AFFFFABRSUUALRRRQAUUUUAFLSUtAwooooAKKBS0AFFFFABRRRQAUtJRQAtFJRSAWikNLQAUUUUALRSUUAFFFFAxaKSigQtFFFABRRSUDFopKWgQtJRRmgYUtJRQAUtFJQAtFFFABRRRQAtFIKWkAtFJRQAUtJRQAUtJRQAtFFFAB3ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAopKKAFopKKAFooooAKKKKACiiigA/GiiigAooooAKKKKACiiigAooooAKKKKACiiigAoozRQAUUUUAFFFFABRRTJXEaFj+AoAhu5MDyx+NVKGJJLHqaBWyVkZt3CkoooEJRRRQAtJRRQAUtFFACUU6igBtFLikxTAKKKKACiiloASkxS0tACUUUUAFLSUtABRRRQI0aWkpaxNRaSiigBH+430P8q5jNdM/3G+h/lXMVpTIkOqxYTJDch5CQuCOBVTNITWjVyLnQG+tP+ev6U3+0LMf8tf0rBzTGNTyIrnZY1GZJrp5IzlSBg1UJpCaaTWiViGxR94fUfzrrh0H0rkFPzD6iuuHT8KzqdCoEV2f9Fm/3DXGg8Cuyu+bWf/cNcbxgVVLYU9xc1ueHADJcDvtH86wCaVJXQ5VmU+oOKuUbqxKlZ3O4Kn0NN6djXEtdT/8APaT/AL6NRmef/nvJ/wB9Go9k+5XtDuc+xozXDCaf/nvL/wB9Guo0NmbTUZ2ZjubknNTKHKrjjK5o0maM0VBYlFBpM0ALSUUmaAFpKKKYgpKKTNAC0lFFABSGiimAUUlFABRRRQAUUlLQIBS0gpaAFrF1W48yXyVPyp19zWhfXH2eAkffbha5/wB+pralHqZzfQKDS0hrYzEpKWigQlFFIaACn28L3FwkKdWPX0Hc1HWxp8PkWpmYfvZxhfZP/r1MnZDirssSlMqkQxFGNqD29fxplFJWRqLSGikoAKKKKACkoooAKSikNAhVBdgqjLE4ArcCCCFLde3LH1NVdLhChrpx04T6+tWCSTk9aym7uxpFWQUmaCaSpKEJqS1XJa4YcD5U/qah2tI6xL1bv6DuauNgYROFUYFD7AhKKSipGKKSjNFAC0lFFABRSUUALRSUtAC0UUUAFFFFABS0lLQAtJRRQMKKKKAFopKKAFopKWgAooooAKKSlpAFLSUtMAooooAKKKKQBRRRQAUUUUALRRRQAUlLSUDCiiloASiiigBaKSloAWkoooAWkopaACikpaACiiigBaKSigBaKTNLQAUUUUgClpKKAFooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigBKKWkoAKKKWgAooooAKKKKAA0UUUAFFFFABzRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABQKKKACj60UCgAooooAKz7iTzH4PyjpVi6k2rsHU9fpVOtIrqRJ9BKWkoNUSFFFJQAUtJS0AFFFFABS0UUAFFFFACUUUUAJRS0UAFJS0UAJS0lFAC0UUUwEopaKACiiigDRzRRRWJoLSUUUwEb7jfQ/yrl66dz8rfQ/yrls1pTImLTTQTVvSlV7wK6hhtPUZrR6K5G5SLUwsOma6vyYf+eUf/fIo8iD/AJ4x/wDfIqPaIrkOTPNNNaWtqsdxGEQKCnYY71mE1ondXIaswU/MPqK6/sPpXHA/Mv1H867DsPpUVOhUCK7OLWb/AHDXGFuBXZXf/HrN/uH+VcX2FVS2FUAmkpcVq6DbwzSzGaMPtUYB6Vo3ZXISu7GRikNdqLKy/wCfWL8qDY2OP+PWL8qz9quxXszhmbHeuq8PEnTI/wDearZ0+xP/AC6RflU8cccSCONAiDoB0pTqKSsOMWmOooNJWRoFGaSimIM0UUmaACikzS0AFJRSE0AFFJRTAKKSigQtFJRQAUUlFAC0CkpaAFozjmiqOpz7I/JU/M/X2FOKu7CbsjOvZ/PnLA/IOFqvilxS11LRWMNxtFLRQAmKbTqSgBKQ0tBpiJ7C2+03ARuI1+Zz6KK1ZX3uTjA6Aeg7Ckgi+y2gjPEsuHk9h2H9aSsW7u5qlZBRSGkzSGFFFJTAKKKKBBSUUE0AFPgiaaVY16sfyplatjF5FuZmH7yThfYVMpWRUVdk8m1QsSfcQYFNpBRmsTQSmsQBntTjSInnShP4By59vSmIlt12RmU8PJ09lp9K7bmz0Ham1JQtFJRQAUtNooAWiikoAWiiikAtFJRQAtFJS0AFFFFAC0UUUALSUUUDFoopKAFozSUUALRSUUALS0lFAC0UUUAFFFFIAooopgLRSUUALRSUtABRRRSAKKKKBhRRRQAZooooEFFLRQAUUUUDCiiigApaQUtABRRRQAUUUUAFFFFABRRRQAUtJRQAtFFFIApaSigBaKKKAEopaKACiiigAoooFABRRRQAUUUUAFFFFABSUuaSgAooooAKWikoAWiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAFFFFABRRRQAUkjhELHoP1pao3Mu99oPyr+tOKuxN2InYsxY9TTaWitTMKKKKAEopaSgAoopaAEoopaAEpaKWgBKKWigYlIRS0GgQlFFLQAlFLRQAlFFLQAlFFFABRRS4oASilopgaFFJRWRoLSGlpKAGv9xv90/yrlc11T/cb6H+VcofatKZnMXNXdI/4/R/ums85FXtGP+nD/dNaS2JW5v0CikNc5qYXiA/6TF/uf1rIJrR8QsftUX+5/WsrJNdMF7qMJPUXd86/UfzrslPA+lcci5Zc+o/nXbbQAPpUVXsVArXX/HtN/uH+Vcb2FdrdAfZJv9w/yriCenNOl1CoOBrb8OEb7j/dH86wc1teG+ZLj/dH86ua90mG50FJS0VzmwmaDSUmaAFpKKQ0CA0lLSUwCiikoAKKSigANJQaSmAtJRSUCCiiigAooooAKKKKYCiikpaAGyOI0Z2PAFYcztLI0jdSauahNuYRKeF5PuapVtBWVzKbuMIpKdQRWhIzFFKRSUAJijFLRigQ0irmmQCSYzSDMUPzH/aPYVVALMFUZJOAPWtlkEESWy/w8uR3bv8Al0qZvSxUV1Edmdi7HJJyTTaSisywNJRRTADTadSUAJmig0lAC0lLR9KBE9jb/aJwp+4vLH2rVlfc+R0HAFNhi+y2wj/5aPy9JWLd3c1SsgpDRSUhjXbapNWYkMUIQ/fb5n/wqKBA8m9vuR8/U1KSSST1NJ9gQUUlFIYUUUUCCiiigYUUUUAFFFFABS0lFAC0UlLmgBaKSlpAApaSgUALRRRQAUUUUAFFFFABRRRQAUtJS0AFFFFAwpRSUUALRSUtABRRRQAUUUUALRRRSAKKKWgBKKKKBhRRRQAUtJS0AFFFFABRRRQAUUlLQAtFJS0AFFFFABRRRQAUUUUgCiiimAUUUUAFFFFAC0tJRSAWkoooAKWkooAWikooAWikpaACikpaAEooooAKKKKAFpKKKAFopKWgBKWkpaACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAo/GikZgilm6CgCK5k2JtB+Zv0FUac7F3LHqabWqVkZt3CiiimIKKKKACiiigAooooAKKKWgBKWkooAWiiigYlFFFAgooooAKKSloAKKKKAEopaSgApaSlFABRRRTA0KSikrIsWikzRQMKpNa2AYh1RWPOC+P61drN1iz86ITxrmSMcj1WqjuSyYWWnHtH/38/8Ar06O1tYD50SqOMbt2RXLeZitLS7uN1eynAMcn3c9j6Vbg11IUkbiuCMggj1FOzmufhlk0u6a3myYGOQfT3H9a3omV0DowZT0I71Mo2KTuQ3FrBc7RNGH29PaoRpdiP8Al3H5mr9JS5mOyKR06x4/cAY9GNXC1IRSYo3ARwHRkblWGDVAaRYAg+T07FjWhRmmm1sJpMpHTbL/AJ9o/wAqlhggtEby0WNOrGppJEjQvIwVR1Jrm9U1I3X7qMFYQe/VvrVJSkJ2Rvfarf8A5+Iv++xSfbLbP/HxF/32K41gKjIGelaeyRHtGdutxC7BVmjZj0AYGpKxtAsPLT7XKuHYYQHsPX8a2aykknZFp3QUUUlIYUGiigBKKDSUAFJRSUwCiikoEFFFJTAKWkozQAUUUUALRSCloAKiupfJiLfxHhfrUuaybubzpSR9wcCqjG7Jk7Ig5JJPJoxRS1uZDTTaeaSgBlJT6Q0ANopafDC08yRJ1Y9fT3ouBd02LYjXjdR8sQP97ufwqT61JIy/LHH/AKuMbV/x/GmGsr31NNtBKSlpKAEoNBpDTAKSiigApKWigQVe0yAPIZ3HyR+vc1TRWkdUQZZjgVtMFhjS3TovU+pqJvSxcV1GsxZix6mkopDWZYGmNu4VRlicAU6pLYcGcj2T/GjYRIQI0WJTwvU+pptHeipGFFJRQAtFJRQAtJRRQAUtJS0AFFJS0AFFFFAwooooAKWkozQAtLSUUgFooooAKKKKACiiigBaKKKACloooGFFFGaACiiigBaSiigBaKSloAKKKKAFopKKQC0UUUALSUUUAFFFFABRRRQAtJRRQMKKWkoAKWkooAWikpaAFopKKAFopKWgAzRRRQAUUUUAFFFFABRRRQAUUUUALRSUUgFooooAKKKKACiiigAooooAWkoooAKKKKAAUtJS0AFFFFABRRSUALRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUcUUAFFFFABRS0lABVK6k3NsB4HX61YuJPLTAPzHpVCriupMn0CiiirICgUUtABRRRQAUlLSUDCloooEJRS0UDEpaKKAEopaSgAooooEFFFFABiilpKACiiigAooFFABRRRQAUUUUwL9JRRWRYUUlFABRmikNAHP6vYeSxuIV/dMfmA/hP8AhWUBg12ZAIKkAg8EHvWFqGlmPM1sC0fUp3X6e1bwn0ZlKPVElvLFqEAtbogTD/VyetVlkutMnMZ4HdT91vcVn78VpQ6hFcRC21AZX+GXuPrTtb0Fc0INWtZQA7eU/o3T86uLIrjKMGHqpzXNXmnSxDzIz5sJ6OvP51RXchyjsv0NL2aew+drc7Yc0uPauOW8vE4FzJ+dBvrtvvXEh/4FS9k+4/aI62R0jGXdVH+0cVl3esW8QIhzK3twPzrBZy3LEsfc5phqo00tyXPsSXV5cXT7pn4HRR0FQZ9adTCDkAck1qjMCa1NI043DLcTLiEHgH+M/wCFSadpDORNdqVXtH3P1rfAAAAGAOABWc59EaRj1YtJS0lYmgUlLSUAFFJRQAUhopKYBSUUGgQlFFJQAtJRRTAKKKSgBaKKKACilpjsEUs3QUAV72Xanlg/M3X6VnVJIzO5Y9TTK3SsjJu4lFFJTEFJS0hoASilpKBCGtOzi8i2Mp/1swwPZP8A69VLOATzhWOI1+Zz6Cr8rmRyxGPQeg9KmTvoXFdRlJS0hqRhSUUUwCkpaQ0hCUlLRimAlFLiprWA3Eyxjp1J9BSv1AuabEI4zdOOTwn+NS9frUkpBIROEUYAplY3vqa2toJSE0GkYgAk9BQABTJIsS/xdT6CrLkZCrwq8CmQL5cRduJJP0FFJ6jQUGiigBKWiigAooooAKKKKQBRRRQAUtJRQAtFFJQAtFFFAwooooAUUUlLQAUUUUgClFJS0AFLSUtABRRRQAUtJS0DCiiigAooooAKKKKACiiloAM0UUUAFFFFAC0lHNFIApaSlFABRRRQAUUUUDClpKKAFpKKKBBRRRQMWikooAWiiigApaSloAKKKKACiiigAooooAKKKKACiiigAooooAKKKKQC0UlFABRRRQAUtJS0AJmlpKWgAopaKAEpaSigBaKSigBaKKKACiiigAooooAKKKKACiiigAooooAKBRRQAUUUUAGKKKKAClFJQKAF60hIAJJ4FFVLuTJ8pTx/FTSuxN2IZHMjlu3amUUVqZi0lLRQAlLSUtABRRS0DEopaKAEpaKKACkpaDQAlFFFABRRRQAlFLRQAlLRRQAUlLSUAFFFFABRRRQIKWiigANFFFMC7RRRWRYlFFFACUhpaQ0wEooNJQIo3um29zlh+7l/vL3+orButOurfcxj3oP415FdXRgVcZtEuKZyNneXFq37mT5e6nkH8KumfT7v/j4hMEh/jj6VtS2VrN9+3Qn1AwarHR7M/dMi/Rs1fPFk8rMttIaQbrW5jlX0Jwaqy6ZfRdbZz7jmt3+x4QcrPID9Kmjs2j6Xs+PSj2gchyjRTp9+GRfqpoUFmCgEse2Oa7MAqMF2f3ak2ru3bRu9cc0e0DkOcg0m6lwXAiX1br+VbNnp9va/Mi7pP77dfw9KuZpKlzbKUUgoooqRiUUUhoAKSikoAKSloNMApKKKBCUUUUAJRRSUwCiiigAoopKACiiigBao3su5vLB4HX61Znk8uMn+I8Csw81pBdSJPoFJS0hqyBKSlNJTAQ0hp1IaAG0U6rmnQgs1xIMpF0H95uwpN2VwSuydY/s9usP/AC0bDSf0FNpzMWYsxyTyTTagsDTTTjTaYBSUUUAFFFFIBKKKKYgJrXtI/s9tkjEsvJ9hVPT7cTTbnH7tOW/wq/I5dy35VnN9C4rqNooNJmoKENESCWXafuJ8zf4Ux22jPU9h61ZVfJiEX8R5c+9DAHYsxNJmkopDFpKKKAFpKKKAClpKKAFopKWgAooopAFFFAoAKKWigBKWkpaACikpaBhRRRQAtJzRS0AFLSUUgFooooAWiiigApRSUtABSUUtAxKWkpaBBRSUUDFooooAWikpaACiiigAooopALRSCloAKKKKACiiigYUUUUCCiiigYUUUUAFFLSUAFLSUtAAKWkpaACiiigAooooAKKKKACiiigAooopAFFFFABRRRQAUUUUAFFFLQAlFLRQAUUUooASlpKKAClpKUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABS5pBR05PSgCOaTy0J7ngVn8nJNSTSeY5PYdKZWsVZGbdxKKWimISilooAKBRRQAUtJS0DCiiigAooooAKKKWgBKKKKACkpaSgApKWigBKXFFFABRS0lACUtFFACGilooEFFFFAAaKDRTAu0UlFZFhRRSGmAVVu721tMfaJ1jJ6A9T+FWv5VycBtI9avP7ZUli37suCVxn/AAxVRjclux0dtdW90pa3mWQDrjqKmzWdb2FoNQS/sZoxFtKukfINaDOKTt0BeZFHcwSyyQxyq0kf31B5Wputc7pDZ13Vcev9a6Fc05RsxJ3HBar3N9ZWrbLi6jR/7pOTTNWuzZadLOn3/ur9T3qjpuk2wtVnuo/OnlG9i56ZoS0uxt9Eakc0M8YkglSRPVTmms2O9cxcr/Y+t27W5Itbo7WTPArpVUk/Q1TjbUSdyOC7guFZoJkkCnBKnODT5Z4oIzLPIscY6sxwBWHp0P2DxDdWR4juF8yP69f8ak8Q5uLux0qM8yPvf6dqfKr2FfQ3k2uiyIwZWGQR3FNldIkaSRgqKMlj0AqUbFUIvCgYH0rD8UzkWcVnHzJcOFx7VEVd2KbsjTguILiPzIJVkTONynIzTya5vRWNjqV3pUjZx88Z9eOf8+1dCrcVUo2Yk7jWuIFnWBpUEzDKoTyRST3EFvH5k8qxpnG5jgZrA1a6jt/EdnNKcRpFkn86bZpLrl2L66GLOJsRRZ6n3quTS7J5uh0qMrqHRgVYZBHcVHJNAkyQPKqyyfdQnlqMkVhajKf+El0z/dP9aUY3G3Y6EgiopZY4kaSVwiKMlj0FSBwaw/FEx+zw2Uf+suJAMew/+vSiruwN2VzXgmiuI/MgkWRCcblPFSVzmgymzvrzSnbIRtyE9/X9MV0IbNOUbME7odSUUUhhRRSE0AIaKKSgApaSloEFITj6UpqtdyYXyx1PWmlcGVp5DI5PYdKjooxWxkJRS0lACYpDS0lABSUppKYCojO6ogyzHAFasgWNEt0OVj6n+83c1DYJ5cbXLfebKx/1NONZt3ZaVkNopaQ0AJRQaQ0wENFFJQIWiiigBKUAkgAZJ6CitDTYQM3Lj5V4X3NJuyuNK7LKoLeBYB948ufem0rEsST1NNNYmgGkJoNMO4lUX77HApgSW6hnMzfdj+77tUhySSeppWCqqxJ91f1NNqdxhR2oNFABRRSUALRSUUALRRRQAZozRRQAUUUUAFLSUUALS0lFIAooooABRS0UAFJS0UDAUtJRQAtFFBoAKKKKQC0tJRQAtLTaWgBaSiigAooooAWkoooAKWkpaBhS0lFAC0UlFAC0UUlAC0UlFADqKSikAtFFFABRRRQMKKKKACiiigAoooxQAUtJSigApaSloAKKKKACiiigAFFFFIAooooAKKKKYBRRRQAUUUUgClpKWgAooooAKKKKAClpKKAFopKKAFpKWigBKKWgUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRzQAUUUUAFVruXH7tT9amlkEaFu/as8kkknqaqK6kyYlLSUtaEBRRRQMKKKKACiijFABS0lQX0zW9nPMmN0aFhnpxQIsUCssXN/B9kkuTA8VwwX92CGUkZH1p9zNeDU1tIHhjUxeYWkUnvj1quULmlRVVrh7Wzee6dHKAsTGMAiqjXepR2/wBskWAxgb2hAO4L9fWkosLmpS1Svrl0t7eS2KbppFUFxkYNENxPHfGyujE7GPzFeMEd8EEGi2gXLlJVa2meaS7U4xFMUXHpjvVe/nuYrm0ggaNDM5Us65AwM0Ja2C5o0VFbC4Ab7RLFIc8eWuMfWpqQxKKWikAUUCigAooopgJRS0lIAoopaYCUopKWgQYooopgWqDRRWZYUUUUCAdayftlnfSz213bojxHAEhGT9DWrVK902zvG3zRfP8A3lOCaqNuonc58rDZ67ax6dISsnEqBsgD0/rXTKhJqCz0yys2LwRYcjG5jk4q4DjpTlK+wkrHPaR5cWu6p5kiICeNxA71vrNCxxHNGx64VgazrjR9PnmeaSFi7ncx3nrS2mlWdpL51vEVfBXO4ninK0tRK60E8QRvc6XLHEMuuHAHfFRaXqUF1ZRZmRJEUK6scEYrUCCqU+j6dPIZHgw55JQ4zQmrWY2ne5jX2NV1qzgtzvjgO53HTrzXUgAHNQ2lrbWaFbeIJnqepP41KxzSlK+iBKxh+J1MRs9Sj4e3kAbHoah0h/t2q3mqsDt/1cQPYVsXVvHcwPBMu6N+CM4ptpZw2kIhgXbGCTgnPWqTSjYTTvckJbtXKzm91LXpXsnQfZAFUv0B/wAc1120VVtLC2svMNuhUyNuckkkmiMkgauc3eWup2d5bapeyxSFXCEx8ce9dWsYwCDweQfam3MEV1A0E6bo26jOKeoCRrGmdqgAZOeKJSugSsc5q9rDceI7SGbmNoxkA4zjNJu/sLVMc/2fcH/vk/8A1v5Vsy2MEt5HeOpM0Y2q2T0+lPubOC6hMNwm+M84quZaJi5epKfmGV5BrB1GInxJpv8Aun+ZrfgjS3hSGMEIgwMnPFRTW8EtzFdOhMsQwpz0/CpjKzG1dEoULXLXH2vUdamnsWQfZSFUv0Hb/GuoYbqrWdlb2aMsCFQzbjk5yacZJaikrnN3UGo2d7b6lfNG+GCFk9Pf8M11oxgEHIPINRXVvDdwGCdd0ZIOM4p0aLFGkaZ2qAoyc8USldAlYfRSUVJQGm0tFACUUUUCFopKKAEdgilj2rOclmLHqanuJNzbR0FQYrWKsQ2NopaDVCG000/FJQA2jFLikNAgp9vC08yxLxnqfQdzUZNaVvH5Ftk/62YZPsvp+NJuw0rj5WVmAQYRRtUe1R0GkqCgpDRRTASkpTSUAJSUtBpiEpaKKAHwxtNKsa9WP5VrybVCwp9xOKhsY/JgM7D534X2FOrKTuzSKsgpKKKkY01LbDajTkfM/wAqew9aiVPOkWIdDyx9BVh2BOBwo4AofYF3EpKKKQwoooNABSUUtACUUtJQAUtJS0AFFFJQAtFFAoAKKKKAClpKKAFooopAFFFFAC0UUUDClpKM0ALRSUtABRRRSAKWkooAWikpaAFopKKAFopKWgAooooAKWkooGLRRRQAUtJQKAFpKWkoAKKKKAFooooAWiiikAUUUUDCiiigAopaKAEpaKKADFFFLQAlLSUUALRRRQAUUUUAFFFFABRRRSAKKKKYBRRR2oAKKKKQBS0UUAFLSUUALSUUtACUtFJQAtFFFABRRRQAlLRRQAUUUUAFFJS0AFFFFABRRRQAUUUUAFFFFABiiiigAooooAKKBUF1JgeWOp6/SmlcGV55PMfj7o6VHRRWpmFFFFABSUtFACUtFFABS0lLSAQ1U1KN5LC5SNSztGQAO5q5RimnYTRSstPtYfJmZJGlRRjzHLbTjsO1QajGraqkstjLcw+Tt+QZw2frWmRQBTvrcVtLFJrVLvS5LeOBrZXUqqOOR71UmbUJrU2YsZEmdfLaQkbAOhOa2xQTQpWG0YupWjvbWVsiu6JKgYqcEKB1zVm2soLZ3kjVjIwwXdixx6ZNXjgmjAp8ztYVupkWktxa3F6GsbmRZJy6siggjH1qXWIjJLYyfZZJ4kcmREGTjFaQFOFDlrcLaWKOm+WEkWKyltRnJV1xuPr1q9S4pDUt3KQUlLRSASgUtFACUUtFACUUUUAJS0UUwCiiigANFFFAi1RRSGoKCikpKAFpKKKACkoopgFJS0lABSUtJQIKDRRQAlJS0hpgFJS0lACUUtJQAlFLSUwEpKdikoASiiigQlFLSUwEooooASkpaKACkopKBBUcz7E46npTyQBkmqcjF2LHp2qooTYyilpMVoQJSGnUUANxTSKeaaaAG0hp1JgkgAcnpTET2MAllLP/AKqP5m9/arcjl3LHqacUEEK246j5pD6t6fhUZrO99S7W0CmmlNIaYDTRRRQAlFFFAgooopgGKmtITPOFP3Byx9qirVjj+z24j/5aPy/t7VMnZFRV2Olfe3HCjgD2qOlpKyLA01iACT0paWFBJLlv9WnLe57CmBJGvlRc/wCsk5PsPSilYlmLHqaSpGFFFJmgBaSiigAooooAKWkozQAtJRRQAtJRRQAUtJS0AFFFFABSikpRQAUUUUgCiiigBaSiloGFFJS0AApaSigBaKSloAKKKKACikpaAFopKWkAtFJS0AFFFH1oAKKKKAClpKKAFooopDCikpaYBRRRQAtFFFIBaKSloAKKKKACiiigYtFFFABRRRQAUUUUAFFFFABS0lFAC0UUUAFFFFIAooooAKOtFJTAWiiikAUopKKYC0UlLSAKKKKAClpKWgBKWkpaACiiigAoopKAFpKWkoAKWkooAWiiigAooooAKKKKACiiigAooooAKKKKACiiloAY7CNCxqgxLMWJ5NSXEm98D7o6VFWkVYhu4dKSloqhCUUtFIApKKKACilooAKKKMUAFMnlMUTOsTyEfwJ1NPpCDTAz4tRaaZolsbkFGCvkD5SfXmpZ7sw3P2dbaaZ9m8+XjgZx3pdLH+m6nz/y2X/0AVV1F411k+ZfNaD7MPmUgbvmPHNVpexPS5eS6BgM0sbwAZJEnUAd6orrULFGe3uI4HICzOuFOen0pLxDc6NNHbXBuiUID5yX56UwalZPBEqfvnbaogC/MD7jtimooTZpyTpHPbwMDumJCkdBgZ5p5lRbtLUg72QyA9sA4rP1NHa+01ElaJiz4dcEj5aZF5sOuRLLcyTk2zHLgDHze1Ll0HfU0Iby3lvprIbhNFjOejfSpIbiKXz9uR5Mhjbd6isowGe61KWA7bqGdWjPr8o+X6Go9Ple9sNTKKUeWVwFPY46UcqDmLLa3b4MgguDbg4M4T5Pr6496sXV7FBCk7ZZHZVBXn73SsmK+tU0kQMcTLF5Rgx827GMYqS8tmt9Bs0nIUxvFvJ7c1XKhczNoZJxVNtSQPJ5VtcTRxHDyRrlQe/1q3BcWk7usFzFIwBJCNnHvWbp+o2djZG2upRFPCzb0bq3PUeuahLyKNQTQvbC6WQeSV3b+2KpLqcZ2M0E8cLnCSuuFPp9KopDcP4cliVCJH3SLH3wTkD8qbd6p50EAsrxd7BU+y+VuOe+c9MVSgJyNie4SGaGOQN+9bardgfQ0R3CSXEsCBiYgNzdgT2+tNv4lls5Vc9FLA+hHINRaRGq6dA4JLSr5jsepY9amytcd3cu0UtJSGFFFApALRRS0AIRRS0UwJ6DRSVIwpKWikAlFFJTAKKKKAEooooAQ0UUlABRRRQISiiimAlFFFAAaSikoAWkooNABSUGkpgFFFIaBBSUtFMBKKKSgBabS0lACUlOpjsFUtTERTtzsH41DQSc5NFaIhiGkpaSgAxSGlpDTAaaQ0tJQITpVzT4wu66ccJwgPdv/rVWjRpZFjQZZjgVoylQFij/ANXGMD3Pc1Mn0HHuRk55NJS02kMDTTS0hpgJRQaSgAooopiDrS0gp6I0jhF5YnAoAtafCGczuPkj/U1ZZi7Fj1NPcLGiwJ91evuajxWLd3c1SsrBSGlNNNIBrtgcDLHgD1NWdvlRrEDz1Y+pqK2UEtOw4ThPc+tOJzyTzQwQuaSiikMKDSUUALSUUGgAooooAKKKKAClpKWgAopKWgApaSloAKKKKAFopKKQC0UUUAFFFFABRRRQAtFJS0AFFFFAwooooAKKKKAFoopKAFopKWgBaM0lFIBaKKKAClpKKAFopKWgApaSloGFFJS0AFFFFAC0UlFIBaKKKAFooooAKWkpaBhRRRQAUtJRQAtFJRQAUtJS0AFFFFABRRRQAUUUUgCiiigAooooAKKKKYBRRS0gCiikoAWikooAWiiigANLSUUALSUUUAFFFFABRRRQAClpKKACiiigBaSiigBRRSUtABRRRQAUUUUAFFFFABUFzJsXYDy36CpXYIpY9BVB2LsWbqaqKuTJjaKWitCRKWiigAooopAJS4oooAKKKKAClpKWgApcUlI7pHG0kjBUUZYnsKAFHBJAAJ68daY8aSHLorH/AGlBqGK+s538uC5jd+u0HmiW7toHCTzpGxGQD6VVmK5ZjRUGFVVHoBilCRqxcRoHPVgozTYpI5IxJG6uh6MpyDVc6lY+d5H2uLzM4257+lKzY9CwwDEEgZHQ46UBFzuwN2MZxzQXVVLMwUAZJPQVHBeWlwxWC4jkYdgeaNRaD9gBJAAJ6+9IEUZwAMnJwKjuLy1t3CTTpGxGQG9Kkhkjmi86ORWi/vA8UagOWOPfv8tC/wDe28/nT22sMMoYehGaghubaZikM6Ow7A025u7a22/aJkj3dMnk0WbGS7I1yUjRSf7qgU1oUdgzIrMOhK5IpYpElQSRuHRujA5BqSgBAuKb5cYcuI0DHqwUZNPpKQARxz0pAAAAAAB0Ap1FMBKKWg0gEooooAKWiigAooopiJqKKKkoKSlpKACkJ+lFQXlrFdw+TOGKZz8rY5oAnyv94fnSEj1X86yToGmf885f+/poGgaZ/wA85f8Av6aq0e5N2apI9V/OkyP7w/Oss6Bph/5ZS/8Af00f2Dpo6RSf9/TRaPcLs1Py/Cg1XtLSGzRo4AwVjk5bPNT0hhmlptLQAUUUlABRRSUAFJQaKYBSUUUAFFFJQIKSlpKACg0UUwEpKWkoASilppoEFV5W3Nx0FSSvgYHU1BVpCY00UtFUSJSGlpKAEpDS0hoASmmnVLbQ+dKFPCDlj6CnsIntF8qEzH78gwnsvc0tPkbexIGB0A9BTKgsSkpaSgQlJS0lMBKQ0ppKYBRRSigQlaVhH5URuGHzNwg/rVS2gM8yoOnVj6CtKZgzYXhF4AqJvoXBdRlJRSGsygpjAuwRfvMcClJqWAbYzMfvPwnsPWnsA59qhY0+4gwPem0lFSMKWkooAWkoNFABRRSUALmikpaACiiigAopKKAFooooAWiijNAC0UlFABS0lFAC0UUUAFLSUUALSUUUgFFLSCigBaKKKACiikoGLRSUtABS0lFABmiiigBaKKKQBRRRQAtFJRQAtLTaWgBaKBRQMKKKKAFopKWgAooooAKWkFLSAWikpaACiiigYtFJRQA6ikpaACkpaSgAooooAWikpaAA0UUUAFFFFIAooooAKKKKYBRRRSAKKKKACiiimAUUUUALRSUUgFopKKAFpRTaWgBaQ0UGgAFFJS0AFFFFABRRRQAUUUUAFFFFAC0UUUAFFFFABmiiop5NiYB+Y9KEgILl97bR90fqahoorVaGYUUtJQAUUUUAGKKWigAooooASilooASilooAKqasT/Zd5/1xb+VW6bPCk8LwyDKOpVhnHFNOzE9jJiguLt7A/YxAkBVzKxGWGOgx61PcvLHrZMFqLhvs3KkgYGfer6ARqqr0UYFIIo/tP2nafN2bM57ZzVcwrGRB5i6LeXMO0PN5kqpH0Q+go+z2h0DGxDD5G7dgdcdc+ua2ILeGAOIk2h2LsM8ZPWqw0qwD7xCcZ3bN52Z/3aamg5SjeK7aRYC4JCO8YmJ9Pf8ASrdxFZrcWzS7Y5A2IdvBJx047VfmjSaNopUDRsMFSOCKqW+m2dtIJY42LgYUu5baPQZ6UlIOUrTef/bf7iKKRvswyJDgAbqs3Z3aZJHdNHbllw5X7qc0XWnW1zMJpBIJAu3KSFePTinJZwrbm22l4iCCHbdkH3NF1oGpQFxLDd2cd7axHJ2wzwnAzjuPpU00Eram09o9vLMIwrwS9QOxB7UtvpVpbzLKiOXX7u9y236A9KnubC1uZFllVhKBtDo5U49MihtdAsxmleS1u/lQeQRKwePOQH74q7UdvDFbxCKFNqDtnP41JUt3ZS2CiilpAJRRRQAUUUUAFFFFABRS0lABRU0YEaGd+g6D1NFF30DTqFFFFIYlJSmkoAKKSigAooooASkpaQ0AFJS0lMQUUUUAJRR3ooAKSiigApDS0lMBKKWkoEFJS0lABSUtJQAUlFFMApDRSUAFNJwCTTqilOflFNIRExycmkp1NNWSJSUppKAENIadTTTEJSU6mmgBDWgieRAI/wCN/mf29BUFnGCxmcZSPt6nsKmZizFieTyalu7GtBKQ0tIaBiUlLSGgQlJS0hpgJRRRTAKWirVjCJJTI/8Aq4+T7mhuyuJK+hagT7Pb4I/eScn2FJSu5dyx70lY+pqFIaKRjgEmgBFTzZBH0Xqx9BU7tubI4A4A9qRVMUW0/wCsflvYelNpbjFopKKQAaDRRQAUlLRQAlLRRQAlFFFABRS0UAFJRRQAtJS0lAC0UlLQAUtJS0AFFFFABSikpaACiiigAopaSkAtFFAoAWiiigApKKKACiiigYUUUUAFLRRQAtFFFIAooooAKKKKAAUtJS0ALRSUtACUtJS0DCiiigBaKKKACiiikAtFFFAC0UUGgAFFFFACiikpaBhRRRQAtFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFIAooopgFFFFABRRRQAUUUUgCiiigAooopgLRRRSAKKKKACiiigAoNFFABRRRQAtJRRQAtFFFABRRRQAjEKCT0FUXYuxY96kuZMtsHQdahrSKtqQ2BpKWimIKKKKACiiigAopaKAEooooABS0lLQAUUUUAFFFVNUyNNuyCQRExBFNagy3inBT6Vz9m9sbmzWw8/zuDNu3bduOc596tak0R1ZUlgmnT7PkJFkkHd14NPl1sLm0NU5oFZ6h/wCzXSyjeFyrbFkyGU/jWbaNDFcWqslzaXOcN5mSJeORnpTUbi5jo6bms4lzrsaknb9mJxnjO6nWysdavRk4EcXGeBS5R3NDNKBms3RlY285Yk/6RIOT71HrNrABHOVbzHmRWIcjjPpmlZXsF9LmttpCpFVprcw2M8FhlHIOzLE4P1NZ1i1vHdW6Fbm1ueQyS5Il46Z6UJXC5tAUUZopDCiiigBKKWigBKKWkoAKWiigAp0aF2Cj8abUkreRDtH+sfr7CgCG6l3OEX7icCioDRWiVkQ3cvUUlBrI0ENJS0lABUc8yQQvNK22NBljjOKfVDXWxpF4f+mf9RTSuxPYfBqdlPKsSSne33Qyld30zUl3e29mEM7Fd5wuFJyfwrFun1DydNkvooltFkQhojls44zmr2tpKLrTPICGUzNtD5wfl71fKrom7sW7S9t7tmWBmJUZOVK/zqvJq1jG5jaRwwJGPLbk/lUts16C/wBrSFf7vlEnPrnNVNVkP2zSv+vg/wDoNCirg27FwX1ubVrvc3kjgnYc/l1qJNX0922rK5bjjym4/Sr/AN7NZthuXV9W57x/yqVbUepbvLq3szGLhypkztAUtnH0qO3vbe6LCBmYrjOUK/zqpq0s39p6V5AQzbpNofOOnfFaFqbpgftiwq2Rt8ok8e+aLKwX1KLa1p6uytMQVJVjsbAP1xV5JFkQOhDKwyCO9c4t1c2tpftHaJJB9ocNKz8LnjkelbWm2/2axgg3h9qD5h0PerlFISbZNcXEVvC88zbY0GScZxUVtqNlcyrFHN+8boGUrn6ZqLXk/wCJNdn/AGR/OoXW+u2s0ezWCOFlcyGQMSAOgx60kk0DbTNS5kjtoXnmcJGgyzHtSK6uquhyrAEH1FZmt3cL3NrZyK7RFvNmCKWO0dBge9SaDMj2bQEktbuU+YYO3qpx9KVvduO+thzatYJIyPPtKnax2NgH64qzNcwxQfaHfMWAdyjPX6Vzov7i1g1Bo7RZIRcOGkL8Ln1HpWzpkDW9jbwFw5VPvDoe/FVKKRKdyMa1pzNtWZ92cY8tuv5VcubiG1UNO+3JwBjJJ9hVLTPl1HVSf+eyf+g1LqCtLc28ltPGl3ErFUkGQyng8f1oaV7DvoT291BcqzQybtpwQRgj6g1JWbZXDSX1xDcWqQ3gRWZkOQ69BWkB60mrME7hQaWkNIYlJSk02gQjtge9Q0rHJpKpCYhpKU0lMQ00lOpKYDTSU6koENoVWdgijJJwKWrdsnlxmc/ebhPYdzQ3YaQ9wEVYU+6nf1Pc0yikpDFpDRSGgQUlLTTTADSUUUAFFFFMQqqWYKoyxOAK1WUQxLAvblj6mobCLYhuWHPRB/WpDknJPJrKTu7GkVZXEoopKkYGnQqGcu33E5+pqNieABlicAVO4CKsK87eWPqaGCGsxZiT1NJRRQMKKKKQBRRRQAUUUUAFFJS0AJS0UUAFFFFABSUtFACUUUtABRRS0AFJS0lAC0UlLQAUUUUAFLSUtAC0UUUgCgUlLTAKKKKQBRRRQAlFFFAC0UCloAKWkooGLRRRSAKKKKACiiigBaSiigBaKKKACiiloAKKKKAClpOlLQMKKKKAFooopAFFFLQAUUUUAApaSloGLRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUlFAC0UUUAFFJRQAtJS0UAFFFFIAooopgFFFFIApaSimAUUUUALRSUtIAooooAKKSloAKWkFLQAUUUUAFRzSbE4+8elPJABJ6CqUjl3LHp2FOKuJuwyilorQgSilooATFLS0lABRRRQAYNJTqKAG0U6koGFFFFABRS0UAFQ3UP2i2mg3bfMQrnGcZqaii4ilJZ5ktJUlKyW425x99ccg0T2c73ou4LoQt5flkGPdkZzV3FFPmYWK72zTWj29zMZGcEF0Gw49qprptwzwC5vjNDCwZU2AEkdMnvWpS0KTQrFO6tHlmjuYJzDcIpUNt3AqexFOsbZrdpZpZjNPKQXcjA46AD0q1RRd7DsZ8VndwF1gvlWNpC+0wgkZOTzmp7+3+1Rom/btkWTOM5x2qxRii/ULENzCtxBJC5YLIMEqcEfSqiWE7TQPdXzTRwHcibAuTjAJPetKko5mhWFxSUtFSMSiiimAUUUtAxKKKKBBRRSqpZgo6mgCSIKAZX+6v86pyO0jl26mp7uQDEKfdXqfU1WpxXUTfQKKKKskuUUUVkaCUh60tJQAtQ3VvHdW8lvLnZIMHacHFTCigDMi0e0jeNnkuZRGQVSWQlQR04qxqFrHeiLzHlQxsWVo2wQasmkqrvcVlsUrWyFszsLi4l3DGJX3Y+lPmtI55IJHLAwvvXB74xzVvBoNHMxWEB21DHDHHcXFwu7fOQXyeOOmKkJooGVL6xivXheR5UeLJRo22kZp1pZLalitxcS7scSvux9KtCii72FZFSKyhihnhALRzMzOGOck9altIVtoI4FZmWMbQWOTip8U08UXbCwy6ijuraS3l3bHGDtODTmPAAHAGKQmkz70AQRW0cd3Ldru82VQrEnjA6AelOEEYu5LsbhJIgRhnggdOPWpadTuFinFY28Uc8YUsk7FnDHOSetTWsKW1vHBGWKRjA3HJxU1NNF2wsQxwRxyzypu3TMGfJ7gY4qG8sYbsoz71kj+5IjbWHrzVsUtCbQrIqWVlBaM7oXeV8bpJG3MfapYYVhDhWdtzFjuOcZ/pUtJQ2AE0lFFACUyQ/wAIp7EAZqA8nJpoGJRRRVEiGkpaQ0AJSUtIaYCUhpabTESQRGaQLnC9WPoKtSMGbgYUcKPQUqL5MOz+N+W9h2FMqL3KENBoNJTAKSijNAhKSlJpKYCUUUYoAKktoTPMsY6HqfQVGa1LeP7Pb5IxLJyfYUpSshxV2PlYEhE4ReAKjoorIsKaadSBS7BB1NMB8I25mPUcJ9fWkpzkEhV+6vApppDCikpaAClpKKQBRRRQAUdqKSgApaTNFAC0UlFAC0UUUAFFJRQAUtJSigAooooAWkoooAKKWigAopKWgAooooAWikooAdSUUUALRRRSAKKKKACiiloAQUtFFABRRS0DEpaSlpAFFFFABRRRQAUUUUALRSUUALS0lFAC0UlFAC0UUUALRSUtAwpaSloAKWkopALRSUUALRRS0DClpBRQAtFFFABRRSUALRSUUALRSUtABRSUUALSUUUALRSUtABRRRQAUUlAoAWiiigAopKKAFooooAKKKSgBaKKKACiiigApaSikAUtJRTAKWiigApaSikAtFFMlcIme56UAQ3Mmf3Y/GoKP50VolYgKWkpaYgooooGFFLRSASjFFLQAlFFFMAooooAKKKKAAUUUtIBKWiq9210sObOOOSXI4kbAxTWomWM0Vl+Zrp/5c7L/v6aUPrv/PpZf9/TT5RXNPFFZvm62P8Al0sv+/ppDLrnazs/+/po5QuaeaWoLU3DQg3UaJLk5VDke3NTUmMDRS0lIYlFLSUwClpKWgAooooAKKKKBBRRRQAhqQuIITJ/G3C0RJvbn7o5Jqpcy+dKWH3Rwv0oSu7BewzPqaWmilrQgWiiigC5RSUVkaBSUtFACUUtVdS8/wCxTC2jZ5mXaoXtnvTQmZljfXM2oFnfNpcF1gHoVP8AWtlcjrWRLpEttZxta3E8stvh44mxgnuB+tbRYFQdpUkZIPaqlboJX6lHUbmZZLaztWCTXBP7wjOxR1P1qFY9Qt7qMfamurdsiTzAAyHsRijUoLhpra8tFDzW5P7snG9T1H1qW1lvrm4jLWZtbdcmTzSCz+gGKeyF1IFmmOtm33nyvs2/b2znrV+7ZorK4kQ4dY2ZT6GqN4txb6x9qispriM2+zMeODn3p9xLcXWmXI+yywysjKsb43E0PWwLS4yxYyLbSNrRd3UMYfl5OOnrS61cyW1oxg/18jCOLH941UsZWjW2RtClSRVVWl2LwccnNS31rLfajApaSK3gUuJE7ueAB+FO1paivpoTaRdSy2m2dszwuYpT6sO/41Sg1G4XWbmK4fdZmURIx/5ZvjIH0NWLOxns9SkCNJNb3CbnkcjKyD1+opsGnPI2qxXMZWK4mBQ+vy9R9DT93UNS3p5klvtQjkcskUqhB/dBHSsy0ledZXn1w27iV1EZ28AHjrVzQo7u3kvDej52dcP/AHwBjNVLOOSBJVm0SSdzK7B9qnIJ460u4F3W5JrezhMMxR2lRDIMZwepquk89tqUFsb8XaShsggZTHfiptZilu7O3jSAn98jOnovfNRJYiz1bzLWAfZZk2ttH+rYdD9DQrW1B7mqGzRSAYpagoSkpaSgAooopgFJRTXOBgdaBDHOT7UynGkqhDaKWkNMApppaQ0CG0UtNNMAqa1jBYyuPkTt6nsKhVWdgijJJwBV1wEVYk+6vf1Pc0n2BDWJYkk5J60lFJSGBpKKQ0wENFFJTEFJS0UAIKWkpVVnYKoyxOBTEWrGESyb3/1acn3PpVqRi7ljTiqwxLAnblj6mmVi3d3NUrKwlBooNIBCalQbIt38T9PYVGih3wfujlvpTnbcxND7AJRRSUDFopKKQC0UlFABRRRQAUUUlMApaKSgQtFFFIYUtFFABSUtJQAUtFFABRSUUAL1ooooAKKKKAFopKWgAoopaACiiigBaKSloAKKKKQBS0lLQAlLSUUDFoooBoAKWkpaACiiikAlKKMUYoAKKKKADNFFLQAYooooASloooAKMUUUALRSUtABRRRQMWjNFFABRRS0AFLSUUgFpeaSigBaKKKACiiigYUUUUCCijNFAwooooEFBoooAKKSigBaKKKACiiigAooooGFFFFABmiiigBaKSigANFFFABS0lFAC0UUlAC0lFLQAUUUUAFLSCloAKKKKQAeBk9KpyOXfPbtUlxJzsB+tQ1cUS2JRS0VQhKWilpAIKWiigBKWiigBaSlooASjFLiigBMUlLRQAc0UUUAFFFFABRRS0AFFFFAAaTFLRQAlFLSUAFFFFABRS0UAJRRRQAUfWiimAlLRRQAUUVKm2NGnfovT3NJgiO7fyohAv3m5c1RFK7F3LscknmkrWKsiG7sWlpKWgQtFFFAFuiiisjQWkpaKAEpaKgupmhiMiwyTEH7kYyTQImJpvWsuHVWmlaNNOuyUYI/yj5D781dvbuGy2B1eSWQ4SOMZZqqzTsF0TgUuT2qnbalDPObd4ZYJwu4RyjBI9R60SX8C6hHYEMJZF3K3b6UcrC6LWaNoNUtSvo9PSNpgzGRwiqvX61JdXkVnC007EKOOBkk+go5WK6Jyo9KTbVGHVY3mjintri2MhwjSrwx9ParN7dJZtCjRSyvMxVFjAJyKLNaBdEwp9VrW9iuJXgMUsMyDcUlXBI9RVSfVljuZ7eOzuZjCQHaNQQOPrRZsLo0ioppA9Kq2t5HdW4niJ2HIO4YKkdQaqrrETBpEtbmS3U4Myp8v1+lPlYXRpnmkxiorm7ggs1vOZISV5T0J6027vIbW4toJM7pzhSOg9zRZgT0lRC4jN21qA29Y97HsATgCpaQBSUGkpgKaSlppoAXOKiJyc05j2ptNCG0UtIaYhKQ0tIaAENNp1GKYhtIaU8U+GLzZAM4UcsfQUAS26+XGZT95uF9h3NFPdtzZAwBwB6CmUhgaSikpgFNNLSUCA0lKaSmAUUGkpiFq/YRbENyw5PCD+tVLeEzzLGOnUn0FacrAkKowi8AVE30LiupGfU9aKKSsygoNFOjA5dhwv6mgBSNibP4jy3+FMoJJJJ60UAFFJRQMWikooAKKKKBBS0lFAxaSlooEFJRRigYUtJRQAoooopALRRRQAUUUUAFJS0lABRRRQAtFJRQAtFJS0ALRRSUALS0gpaACjNFFAC0UlLSAKMUUUALSUUUDFopKWgAoooxQAtFJS0gAUtJRQAUUUUwFopKWkAUUUUAFFFFABRRRQAUUUUALRSUtAxaKKKACiiigBaKKKAFopKWkAtFJmigBaSlpKAClpKWgAooooAKSiigAopaSgApaSigBaKSloAKKSloAKKKKACiiigAooooAKKKKBhRRRQIKDRRQMKKKKAClpKKAFooooAKWkpaACmSvsTPftTyQOtU5W3vnt2ppXE2M9zS0UVZIUUUYpALRRRQAUUUUAFLRRQAlLRS0AJRS0UAJSUtGKAEooooAKKKKACloooAKKKKACiiigBKKWigBKQUtFABRRS0AJRRRQAUUUUAJS0UUxDo1LsFHeobyUMwiT7icfU1PK3kQcf6yTp7Cs+nFXdxN9Aoopa0JClpKXvSAKKKKALlLSUVkaBRRRQAtJiiigDM0s7b3Vx/08/+yiq9y/l+IIZJjtR7cpGx6bs8itZIYo3kdECtK25yP4j602e2guI/LniWROuGFWpK9yLOxj3TCbWtPWEhniDtIQc7VI71DqNu8uqu8efOithLH9Q1bltaW1qCLeFYweuO9PEUYm88IPN27d3t6VXPbYOU56/k/tISX2P3MJijjH+0SC3+FaeuLGkthNJ/qI7gGQ9hxwT+NXBbW4g8gQqIt27b2znOakkVZFZJFDK3BBGQaXMFjI12SN7eKCNleaSVPLAOT160utfaDe6X9nKCUSttLglc474q7BYWds/mQW6I/qBzUzxxu8cjoGaM5Qn+E0cyWwWZlact1LqlzLeuguYkEYRBhdh53D1qAi/jvtXlsjCdrKSrqST8vat0pGZhPsHmhdu7vj0pixRpJJIigPIQXP8AeNHMHKZNvGreHpTbyGSSZHckDHzHqMVLZXdqmlRSCRBEkQDAnpxyMVdgtooAywoEVmLED1NRNplg0vmtaRmTOc4707rqKz6FLT42l8Om3kUjzEYqD2GciqQL6nb3FyAd1vboif74OT/KukKj0qKGCG3RkhjVFYkkAdSetHN1DlM7SCbme8vu0rqq/QD/ABNaikEZVge3Bot0it4hFDGEQdAKbHFFCGWJAgZixA7k96Td2NKw6kpTSUhhTWOBTqjY5NCEJSUtJVCCm0tJQAUhoooASiikNMQGrgXyYhH/ABty/wDQVDbIMmVhlV6e5qQkkkk80mNDaQ0ppKAEpKU0lMANNpTSUCA0lLRTASilxVqxhEkhkf8A1acn3NDdlcEr6FmCP7Pb/wDTSTk+wpKc7F2LHvTax9TQKSlpKADBJAA5NPfAwg6L+poX5E3/AMR4X/GmUAFFFFABRRSUAFFGKKACiiigAooooAWikpaBhSUtFACUtFFABRRRSAKKKSmAtFFFIAopRRQAlFLRigBKKKKAClpKWgA60UUUAApaSloAKWkpaACiiikAUtFFABRRRQAUtFFAwpaBRQAUlLSUgFooooAKKKKYC0UUlIBaKKKACiiigAooooAKKKKAFooooGFLSUUAFLSUtABS0lLzQAUUUUgFooooAKKKKAAUtJRQAtFJRQAUUUUAFFFFABRSUUALRRSUALRRRQAUUUUAFFFFABS0lFABS0UlAC0UlLmgAooooAKKKKACiiigYtFFBoEFFFNkcIue/agZHcP/AAD8ago6nJoq0rEC0lFLQAlLSUtABRRRQAUtJS0AFFFLQAUUUUABooooASilpKACkpaKAEpcUUUAFFFFABRRRQAUUUUAFJS0lAC0lFLQAUlLRQAlFFFABRRS0AIakiUEl34ReSaYASQAOTTb2QIot0Puxp76B5leeUyyFz07D0FR0lKK1MxaTNFFAC0tNpaQC0U1jgZop2AvZooorI0CiiikAUlFBoEJS0lFMApKU0lABQaKDQAlJS0UAFJS0lABSUtJTAQ0lLSUCEopaSgApKWmk8UwGse1MNKaSmISilpKYCUhpaQ0CEpKWigBKEVncIoyScCkq1Cvlx7z99xhfYetDBDn2qBGn3V/U+tMpaSgYhpDS000CDNJRQaYCGkpaTFAgooopgKqs7BVGSTgVqsFijWBO33j6mobGMRxm4YcnhP8aeeTkms5O7sXFWQUlFFQUJTlXccdu9NNPPyrt7nk0xDXO5s4wOwpKKKACiiigAopKWgBKWkooAKKKKACiiigYtFJRQAtFJRQIWiiigAooopDCiiimAlLRRQAtFFFIAooooASiikoAWiiigAoFFFAC0tNpQaAFoFFFAC0UUUAFLSUtIAooooAWiiigYUUUtABSUtJSAKWkpSaACiiigAooopgLRRRSAKKKKACilpKACiiloASiiigApaSloGLRSUtAC0UUUAFLRRSAKKKKACiiigAoopKAFopKKAFooooAKKKSgBaSikoAWiiigBaKSigApaSigBaKKKACiiigAoopaACkooxQAUtNpaAFopKKAFooooAKWkpc0AHTk1Ukfe2e3apJ3/gH41DVJdRNhRRRVCFooopAFFLRQAlLRSUALRRRQAUtFFABRRRQAUUUYoAKKKKACkpaKACiiigApKWigBKKWigApKWigBKSlNJQAUUUUAFFFFABRRRQAUtJT4l3uF7dTQA4MIYjM3U8KKzWJYlick9amu5hLJhfuLwtQVpFWV2RJhRRRVEhS0UUDAUtIKRjgUANc5OO1FMoqkSalJRRWBsFFFFIQVBdfaPK/0Xy/NyP9Z0xU1FNAZ3/E59LP8AWmn+2uws/wBa06KfN5CsZY/tr0s/1p+NZ9LL9a0cUU+byCxUtPt+5vtgg24+Xys5z71apaSpYxKKKQ0AFFJRmgANJS0hpgFJRSUCCkpaSgApjHNOPSmU0IKSiimAlIaU0lMBKQ0tNoEFFJR1OKYh8MfmSfN9wcsfap3bcxPT29KcVEaCIdern39KZUlCUUUUwG0hpxptAhKSlpKYBRRiimIKktojNMEHTqx9BUdaUMf2eDH/AC0flvYelTJ2Q4q7HyMCQFGFXgCmUUlZmgUUUuM8UhAuPvEcD9abnJyetKx5wOgpKYwooooEFJS0UAJRRRQAUUUUDCiiigQUUUUAFFFFAwopaSgBaKKKBBSUtJQAopaSikAtJRRTGFLSUUALRRSUgCilxSYoAKKWigBBS0UUAJS0UtABRmiigApaSloAKWkpRSAKWikoAWiiigBaKSigYtJS0lABRRRSAKKKKYC0UlFACilpBRSAKWkooAWikooAWikpaACiiigApaSigBaKKKBi0tNpaAFopKWgA6UtJRQAUUUUgCiiimAUUUUAFFFFIAoooNMApKKKACloFFABRRRQAUUUUgCiiimAUUUUALRSUUALRSZooAWikpaQCUuaKKAClpKWgApsjbEz37U6qsj72z27U0rg2M6nJ60tFLVkCUUtFIYUUUUAFLSUUALRSUtABS4oFFABS0UUAFFFFIAoopaAEopaSgAooooASilpKYCUtFFABRRRQAUlFFABSUUUAFFLSUAFFFLQAlLSUtACUXL+TD5an95J94+gqSPaoaV/urz9TWfK7SSM7dTVRV2JuyGdKWkpa0MwpaSigBaKKKBhUbHJ9qexwPrTKBMQ0UpoqhGjRRRWBsFFFJQIKKKKQBRRRTAKKSigApDS0lABSUUGgBKSlpKYBSUUUCENJS0lABRRTSaYCGkpaSmIKSlpDQAlIaKKYhKbSmkoAQ1PbptBmYdOF+vrUUaGRwo79T6CrLkcBR8qjAofYaG55oNBopANopaSmAlIaDRQISkpaSmIKKM1NbQmeTGcIOWb0ovbUa1JbKEEmeQfIvT3NWGYsxY9TTnYYCoMIvAFMrJu7uWlYKKM0lIYU7oM9zSDmkJyaAEooopgFFFFAgooooAKKSigAooooAKKKKACiiigAooooAKDRRQAUUUUDFoxSUtAAKKKKACiilpAJRRRQAUUUUwFooooAKKBRSAWkoooAWikpaACiiigApaSigBaBSUtAC0UUGkAUopKKAFopKKAFopKWgYUUUlABS0lLQAUUUUAFLSUUALRSUUALRRRSAKWkooAWikpaACiiigAooooAWlptKKBi0UlFAC0UlLQAUtJRQAtFJRQAtFJRQAtFFFACUUUUAFFLRQAUUlLQAUUUUAFFFFIAooooAKKKKYBSUtFABSUUUALRRRQAtFJS0gFopKR22Lnv2oAjmf+AfjUNHU5PWirSsTcKWkpaAClpKWgBKKWigApKWigAopaKACloooAKKKKQBRRRQAUtJS0AFJS0lABRRRQAUlLSYpgFFLRQAlLRSUAFFFFACUUUUAFFFFABRSUtABSqpZgo6mkp7v9nhMn8bcKKAK99IMiBPup19zVSg9c0VqlZWM27sKKKKYhRRSUUALRTc012wMUABOTmim5ozTEONFMJopgatFFFc5sJRS0hpgFFFFIAoopKYgooooAKQ0UhoAKQ0tJTASilpDQAUlFIaBBSZpaQ0wEJxTaU9aKBCUlLTaYBmkoooASiikNMQhpDSmnwoGbew+Vf1PpQBLGvlx8/ffr7CihiSST1NJSGFJS5pDTATNJmiigQU2nZpKAEoNGKmt7ZpjnO2MdWobtqCVxsELTPtXgDq3pV/5UQRRjCDv6mlyiIIohtT9TTKzbuWlYKKKKQwpKWk96AA+lFJS0wEopTRSEJRRRTGFFFJQIKWkooAKKKKACiiigAooooAKKKKBhRRmigAooooAKWkpaACiiigAopaSgAooooAKKKKACiiigAooopAFLSUtABRRRQAUUUUAFLRRQAtFJRQAtLTaWgApRSUtIAooooAKKKKACiiigApaSigYtFJS0AFFFFABS0lFABS0lLQAUUUUgClpKKAFooooAKKKKAClFJSigAzRRSUALRRRQMKKKKAFopKWgAooooAKKKKAClpKKAFopKWgAopKKAFpKKKACiikoAWikpaACiiigAooFFABRRRQAUtJS0AFLSUtAB05qs772z2HSnzP/AAD8aiAqkhNhRRRTEFLRS0gCiiigAooooAKBS0UAFLSUUALRRRSAKKKKACiilxQAUUUUAFJS0UAJRRRQAUUUUAFFFJQAUUUUwCiiigApKKKACiiigAoooHoKAJIlDNz90ck1QuZjNKW/hHCj2qzeyCKMW6n5m5aqGauC6kyfQWkozSZrQgXNJmgmmk0AOoJphNN3UWAeW4z2qEvk5pJX/hFR5ppCuS7qC1RZozTEPLUVHmimBv0UtFcpuJRRRQIKSlpKACkNLSUAFJS0UwEpKWigBKSlpDQAUhoopgJRRSUCCg0UlMBppKU0lMQlJS0lACYpKU0hoAKSjNApiEAJIUDk8CrJwoEY5A6n1NESbV8w/ePC/wCNNIpDFpDRmigBtFLSUwEoNFOSN3OEQt9KBDKACxAUEk9hVxLPHMzhR6Dk1OhjiGIU2/7R61LmuhSj3IYrQJ81wceiDr+NTO+4BQNqjoBTSSTk9aKh67lLQSiiigApKKKACg0UUAFFFFACUUUUAFFFFMQlFLRQMSloooEJRRRQAUUUUAFFFFAwooooEFFFFABRRS0DCikNFAC0UlFAhaSiloAKKKKBhRRRSAKKKKACiiigAooopgFFFFIApaSgUALRRSigAooooAKWkooAWikpaACiiigApaSloAWikopAFFFFABRRRmgApaSloGFFFFABS0lLQAUUUUgCiiigAooooAWiikoAWiiloAKKKKAEpaSigBaKKKACiikoAWiiigBaKBRQMKKKKACiiigAooooAKSlpKACilooASiiigQUUUUAFLSUUDFooooAKWkooAWkdtq5/KioJG3HjoOlNITGdTmlooqhBS0lLSAKKKKAFooooAKWiigAopaKQBRRRQAUUUtAxKWiigAooooAKSiloAKKKKACkpaSgQUUUUAFJRRTAKKKKACiiigANJS0lABRRSUAFPVlijad+i/dHqaail2CjvVTUJw8giQ/u4+PqaaV3YTdlcgdy7l2OSTk0maZmjNbGY7NGaYTTSaAHk03dTCaTdTEOJppbApM0hoAYc5zRSkUlMQUUUUwEopaKBHRbW/un8qUq390/lUmT6mjJ9a5LnSRbH/un8qNj/3T+VSZPqaMn1ouIj2N/dP5UbH/ALhp+T60mfc0AN8t/wC4aTy5P7hpxz6mm8+tMA8qT+6aPKk/u0nNIfrQA7yZP7v60nlP6D86Q02gB/lP7fnR5L/7P50ykpiJPJf1X/vqk8h/VP8Avqo6SizC5J5Df3k/76o8hv76f99UykNGoEhgP/PWP/vqkNuf+esf51Fikp2YEv2f/ptH+dH2Yf8APeL86ipKLMRN9mH/AD3i/Ok+zr/z8RfnUJpuKdn3DQn+zr/z8RUn2ZP+fmOoaKLPuIn+yx97qP8AKgWaFuLhCB1wKr5Oa0IYxGmCPmPWk7rqNWYnkAnmZQOwx0pfsqf89h+VPxSVN2VZDDaJ/wA9h+VN+yL/AM9x+VS0UXYWRELRM8zj8BThb246u7fSnYpKLvuFkOC26/dhB92OacZHPAO0egqOjNAAaKKKACiikoAKKKKAEoopaAEooooAKSlooASiiigApKKKYgzRRRQAtFJRQAZooooAKKWigBKKWigYlFLRQAlLRRQISiiigAooooAKKWkoAWikooAWiiigAooNJQAtFJS0DCiiikAUUUUAFFFFABRRRQAUooooAKWkooAWikooAXNKKSigBaKKKACiiigBaKKSgBaKSigBaKKKQBS0lLQAUUUUDCgUUUCFoopKBi0UlLQAUUtJSAKWkooAKWkooAXNFHWigAooooAKKKWgBKKWigAooooABS0lFAC0lFFABRRRQAUuaSigAooooAKKKKACiiigAooooAKWkpaBhRRS0AJS0U122rmgBkr/AMI/GoqTknJpapIkKKKWmAlLzRRSAKKKKAFFFFFAC0UUUALRRRSAKWkpaACiiigYUUUUAFFFFABRRRQAUUUUAFJS0lAgoopDQAUUmaTNMBaKbmjNADqSkzSZoAdS03NLQAUhNLTokDMS3CKMk0AMnl+zW+4f62ThfYetZGamu5jPO0n8PRR6CoDWsVZGcndi5pM0hpM1ZI7NNJozSHmgBM0UUUALTacaQ0wEpppSabmgQpoopKAA0UGimI6eiig1yHSJRRRQISiiigBDTaU0lMBKQ0tIaAENJS0lMQGkpTSUwEoopKBBSUpptABQaWkNMBKbmlNJQISiiimAhpKU0qIXcKO9AE1pHlvMPQdPrVukVQqhV6ClrNu5aVgooopAJRRSUAFFFFMBKKWkpAFFFFMBDRRRQAGiikoAKKKKACg0UUAFFFFACUlLRQAlFFFMAooooEFFLRQAmKKKKBhS0lLQAUGiigAooooAKKKKACkpaKBCUuKKKACiiigBKWg0lAxaKSgUCFpKKKBhS0lFAC0UUUAFFFFIAooooAKKKKAFoopKAFoopKAFxS4oooAKKKWgAozRSUALRRRQAUUUUAFFFFAC0UlLQAUtJRSAWiiigAopKWgAopKWgYUUUUCFooooGFFFFABRRQKAFFLSUUgFopKKACloooAKKKSgBaSlpKAFoopKAFopKWgAopKKACiiigBaKSigBaKSigAooooAKUUlFADqKSloGFLSUtACVBI25uOg6U6VsDaO/WoqpEsKWiimAUtJS0hB1oopaBhRRQaACiiigBaWkpRSAKKKKAFopKWgAooooGFFFFIAooopgFFJS0AFFFJQAtJS0lAgpppTTDTACabmkJpM0AOzSZphamlqYiTNANQ7qUNzQBODS5qMNS5pAP5JwKZfyeVGLZT8x5c/0qRHWGN7h+i8KPU1kvIzszscsxyaqKuxN2QGmmlzSGtTMaaSg0maYC0ZpM0lAC5opKTNMBaQmkJppNAhc02gmm5pgPzRmm5paBCk0U00UAdVSUGkrkOkKKKKBBSUtIaAENJS0hpgJSUtIaYCUhpaSgQhpDSmkNMApDRRQISkpaSgApKXFJTAQ0lKaQ0CEoopKACrttHtTcerfyqC3j8x8kfKOtXqmT6FJCUUGipKEoopKBBSUtJQAUUUUAFJS0lMAooooAKKKSgAoopKACiiigApM0UUAFFFFAgooopjCkpaKBCUUtFACUUuKSgAopaKBhRSUtAgooooAKKKKACiiigApKWigAoxQKKACiikoAKKKKACiiigAooNFABRRRQAUtIKWgYUUUUhBRRRQAUUUUDClpKKYBRRRSAWlpKWgApaSloASiiigApaSloAKKSigBaKKKAClpKKAFopAaXNIAooNJQAtFJS0AFLSUtABRSUtAwooooAWkoooAKKKKAFooooAUUUUlIBaWkooAWikooAWikozQAtJRRQAUUlFAC0UUUAFFFFABS0lFAC0UlFAC0lFFABS0lLQAUUlLQAtIzbRmjNQSNuPsKaQNjScnJ70UUVRIZooooGKKWkopALRRRQAUtFFABRRS0AApaSlpAFFFFABRRS0AJS0UlAC0UUUAFJS0UDEpaSigQUUUmaAFoNFIaAENRsaeaic0wGMaYWoY0yqEPzTTRmkJoEFKKbSigB+akiVpHCDqahzUtxN9itNw/183C/7I9aAK+p3AeUQR/6uLj6mqYaoQacDWqVlYzbuS5pM00GlFMApKWkNABSUtJigQUhopDTAQmmk0pppNACZpM0hNJmmIdS5puaKAHGikopgdXRRRXGdIUUUhoEFJSmm0wCg0UUAJSGlpDTEJSUtJQAhpKWkpgJSUtIaBBRRSUAFJS0hpgIaSlNNoELSKCWAHU0VatY+PMb8KHoNE8aCNAo/GnUUlZFgaSiloASiiimAlFKaSgQUlFFMBMUUtJQAUUUlABRRRQAlFLSUAFFFFABSUppKBBRRRQMKDRRQAUUUUAJRRRTAWikpaACkpaSgAooooAKWkooELRRRQAUUUUAFFFFABRSUtABRRRQAUlLRQAmKKWigAooooASilooAKKKKBhRRRQAUUUfhSASlFFFMAoopaQCUUUtMAooopALRSUCgBaKKKACkpaKAClpKKAFpKKKAClpKKAFpKWigAopKWgApaSikAtFFFABRRRQAtFFHSgYUUUUCCiiigYUtFJQAtApKWgAooooAKKKKAFopKWkAUUlFABRS0UAJS0lFAC0lFLQAUUlFAC0maWigAopKKAFooooAKKWmsdozQAyVv4RUdByTk0lWhC0UlLQIKWkpaACiiigYtFFFIAooooAWiiigBaKKWkAUUlFACiiikzQAtFNzRmgB1FJmjNADqSiigApKWkoAKSiimAUGikNACNULVIajYUxERptOIpDTEJSGlNJQAlLSUDJIA5J4ApgWLaNXYvIcRRjcxrJvblrq4aU8L0Uegq9qswhiWxjOT96Uj+VZWKqC6kyfQBThQBSgVZI4U6miloAWlpKKAFpKUGimA2mmnmmNQIY1Rk05jUZpiDNJSGm5oAeDTxUYqQdKACilNFMDq6KKK4zoCkoopgIaSlNFACGig0lABSUtIaYhKQ0tNNACmm0tJTASkNLSUCEooopgFIaKSgAzTaWkNAh8UZkcL271oAADA6Co4I/LTn7x61JUSdy0goopKkYcUUUUwEopaSgBKKWkNAhKKKKACkpaSmAUlLSUAFFFFABQaKSgAooooAKDRRQISiiigYUUUUAFFLSUAFFFFABSUUtMBKKWigQlLSUtABSUtFAwooNFACUtFFACUUUUAFLSUtAgooooAKSiigBaKKKACiiigApKWigBKWiloASiiikAUUUUDCilooAKKPpRQAUUUUAFLSUUAFFFFABS0lFAC0lFFAC0lFFABS0UUAFFJRQAtLSUUALRRRQAUUUUAFLSUtABS0lFIBaKSigApaKKACiiigBaKSigBaSiigYUUUUAFLSUUALSUUUAKKWkooAKKKKQBRSUtABRRRTAKKKKQAaKKKACiilpgFFFFIAqF23H2FPlbA2ioqpITCikpaYgooooAKKWjFIYUUUtABRRRQAUUUtABRSUZoAWikJpM0AOpM0wtSbqAJM0hNR7qQtQBJmjdUO+l3U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AHJNIB0KqMyPwicmsy5laeZpW79B6CrmoShQLZDwOXPqaoGrgupMn0GdKA1NeoScGtCC0HFO3VUD04ygDrRYLlkNT1wao+d71NFMPWlYdy1toxTFkBp4INIYlNNONJTENNNNPppoAbUUpwKmxUMqkimJmbcSlc1SknzVu6jzmsx0K5rREMbJITUJaldqizVCJQaniqoCasRNSAvxsRViOX3qpGeKsRKSc1Iy6j5FFJGvFFIo6eilpK5jcKKKKACkpaQ0AJSUtJQAUhpaSmISkpTSUAJSUtJTASg0UKpdgo6mgRPbR7m3H7o/nVykRAiBR2pazbuaJCUUUUhiUUtJQISiiimAlFLSUCCiiigBKKKKAEooooASiiimAUUUUAJRRRQITFJTqSmAlFLSUAFFFFAAaSlooASloooASilpKACijFFACUUtFACUUtJQAUUUtACUUUtACUUtJQAUUUUAFFFFABRRRQAmKKKKYAKKWigAooopAJRQaKACiiigAooopgFFFFABRRRQAUUUUAFFFFABRRRSABS0lFABS0lFAC0UUtACUUUUwCiijvSAKKKKYBRRRSAKKKSgBaKKKACiiigAopKKAFFFFFABRRRQAtFFFABS0lFABRRRQAtFIKWgA5ooooAWikpaACkpaMUAFFFFIAooooAKKWkxQAUlLRTASlpKKAFpKKCQBmgBrtjgVHSEknNGadhC0UmaCaAA02gmkJoADQTSUmaYBmjNJRmgQ7NBJpmaM0AOLU3dTSaTNAEoanA1CDTwaAHGkNOoxQBHijFOIooAbtoIp9IaAImFRsKlaomNMREajNSORURNMQUlNJoBpgSLUyioENWI6TGOC0bakUUpUUgIdtGypDxSUANC09RSqKeBQA9RUoAqEHFSBqTGSDFGaZupC1IY8tUgkWCBrhup4QepqtGplkCDj1PoKo316s0+2M/uo/lX396ajd2E3ZXFZizFick8k00tUIlB70jyVrYzFkeqsklLK9VHY1SQrkxlqN5veq7MaiZs96dhXJ/POasRTe9Z4NPViO9FgNiOb3q3HJmsKOUir0M3vUtFJmoGBFLVVJM1Mr5qbDuPxSGnA5pDQMbSMMinUlAilPHkVQmt89q2ygNQvEKpMTRzU1sRniqrRMprpJYQRVOS2BPSrTIaMZVYngVahgc1oRWgz0q7FbAdqLgkUobcjtVtIsdquJEBSsoqLlWIFwKKHU80UwOiopaK5jcSig0UAJSUtJQAUUGkoASiikpiCkpaQ0AJSGlpKYCVctYtq726np9Kggj8x+fujk1fqZPoOKEopaKgsSkpTSUAFJS0hoEJRRRTAKSlooASiiigApKWkoAKSlooEJSUtFACUlLRTASilpKACiiigQlFLRQAmKKKKADFJS0UwEopaKAEoxS0UAJRS0lACUUtJQAUYpaSgAopaKAEopaSgAooooAKKKKAEopaKACiiigBKXFFFACUYoooAKKKKACiiigBKKWkoAKKKKYAKKKKQC0lFFABRRRTAKKKKACiiigAoopaQCUUtFABRQKKAClpKKAFzSUUtACUUUUAFFBoxQAlKKSloAKKKKACiikzQAtJRRQAUUUUwCloooAKWkopALRmkooAWiiigAozRRQAUtJRQAtGaKKAFFFIKWgAooopAFFFFABRRRTAKKKSgBaQ0UGgBKidsn2p0hwMDqaipoTFzRSCimAtIaKbQAE03JpTRigBKKKKBCUGikNAAabmlNNxTAUUuKaOKeKBABTgKMU4UhijilpKaTigBxpM0wtQDTAfRSZozSAawqu9WGNQPzTQis7YqFmqd0NQshqhDM0oNGw05UNMQqZNWY6jRasIuKkZKlSEUwDFPpFETUzmpyuaQqBRcBEp1NPFN30APzRn3qMvSbqQEpfFNMlQs9OttgL3ExxBCNzE9/amA7UZ/sVkIwcXFwOfVVrnvMIpLy9e8upJ5OCx4H90dhUanJrWMeVamUpXZZEhqQMTUCiniqEPbmoHFTZ4qNqAKr8VCTViQVWemIA1Sg1WLYNPR6ALA61MjFahjOanCk9BmkMkFwR3q1DPnvVE20jdsVZtoCvWk7DVzTibIqXFRQjAqUmoKGNxUZentUDmmgHmQVG8o9aifJ6VXkDU0iblguDTQRnk1nyysneqhviGwTVWFc6GPbU64rBtrzcQM1rwSbgDUtWGmWaOKO1NLUhiOooprtRTQG5RQaK5zYKKKKAENJSmigBDSGnGm0AJSGnU2mIKKKKAGmjBJAHWlNWLWP8A5aH8KG7ATRJ5aBe/en0UVkaBSUtJTADSUtJQAYopaSgQlBoooASilpKYBSUtGKAEopaSgBKKWigBKSlooEJSU6koASilopgJSUtFACUUtJQAUUUUCEopaKAE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';

        // text tlacidla -> ikona a podnadpis (hlada sa podla zaciatku textu)
        const KARTY = [
            ['výroba',        '▶', 'aktívna'],
            ['upinanie',      '🔧', 'príprava'],
            ['upínanie',      '🔧', 'príprava'],
            ['upratovanie',   '🧹', 'stola'],
            ['meranie',       '📐', 'v procese'],
            ['programovanie', '💻', 'CNC'],
            ['chyba',         '⚠', 'prestoj'],
            ['údržba',        '🛠', 'stroja'],
            ['udrzba',        '🛠', 'stroja'],
            ['stretnutia',    '👥', ''],
            ['prestávka',     '☕', ''],
            ['čakanie',       '⏳', ''],
        ];

        function kartaPre(text) {
            const t = String(text || '').toLowerCase();
            for (const [k, ik, pod] of KARTY) {
                if (t.indexOf(k) !== -1) return { ikona: ik, pod };
            }
            return { ikona: '●', pod: '' };
        }

        function injectStyles() {
            if (document.getElementById(STYLE_ID)) return;
            const st = document.createElement('style');
            st.id = STYLE_ID;
            st.textContent = `
/* ---------- pozadie a zakladne karty ---------- */
/* Pozadie HF Slovakia (navrh 3 z 2026-09-16) je vlozene priamo v skripte, aby
   nezaviselo od siete; svetlomodry prechod pod nim je zaloha, keby sa
   obrazok z akehokolvek dovodu nevykreslil. Obrazok je len na <body>,
   vnutorne kontajnery appky su priehladne - inak by sa kreslil viackrat. */
body.${BODY_CLASS} { background:url(${POZADIE}) center / cover no-repeat fixed,
  linear-gradient(135deg,#eaf2fb 0%,#dbe8f7 55%,#e9f1fb 100%) !important; }
body.${BODY_CLASS} .sapUiBody, body.${BODY_CLASS} .sapMShell, body.${BODY_CLASS} .sapMPage,
body.${BODY_CLASS} .sapMPageBgStandard, body.${BODY_CLASS} .sapMPageBgSolid,
body.${BODY_CLASS} .sapMNav, body.${BODY_CLASS} .sapMApp { background:transparent !important; }
body.${BODY_CLASS} .sapMPanel { background:#fff !important; border:1px solid #e3ebf5 !important;
  border-radius:18px !important; box-shadow:0 4px 18px rgba(16,36,63,.08) !important;
  margin:10px 12px !important; box-sizing:border-box !important; }
body.${BODY_CLASS} .sapMPanelHdr, body.${BODY_CLASS} .sapMPanelHeaderTB {
  background:transparent !important; border:0 !important; }
body.${BODY_CLASS} .sapMPanelHdr .sapMTitle, body.${BODY_CLASS} .sapMPanelHdr .sapMText {
  font-size:12px !important; font-weight:800 !important; letter-spacing:.14em !important;
  text-transform:uppercase !important; color:#4a6285 !important; }

/* ---------- horny pruh s logom HF Slovakia ---------- */
#${HEADER_ID} { display:flex; align-items:center; gap:12px; padding:0 14px 0 4px; flex:0 0 auto; }
#${HEADER_ID} img { height:34px; display:block; }
#${HEADER_ID} .nd-h-txt { line-height:1.15; }
#${HEADER_ID} .nd-h-n { display:block; font-size:16px; font-weight:800; color:#13315c; letter-spacing:.02em; }
#${HEADER_ID} .nd-h-s { display:block; font-size:10.5px; color:#7d8ea8; }

/* ---------- stavove tlacidla ako karty s ikonou ----------
   Vnutro tlacidla (UI5 <bdi id="...-BDI-content">) sa NESMIE menit - UI5 ho
   potrebuje pri stlaceni a v 2.1.0 to tlacidla odstavilo. Ikona a podnadpis
   su preto len pseudo-prvky ::before / ::after na obale .sapMBtnInner a text
   beru z data-* atributov, ktore na obal doplna JS. Skutocny text tlacidla
   zostava netknuty v .sapMBtnContent. */
body.${BODY_CLASS} .statusBtn { border-radius:16px !important; border:0 !important;
  box-shadow:0 6px 16px rgba(16,36,63,.18) !important; min-width:190px !important; }
body.${BODY_CLASS} .statusBtn .sapMBtnInner { padding:12px 18px 12px 14px !important; border-radius:16px !important;
  display:grid !important; grid-template-columns:40px minmax(0,1fr) !important;
  grid-template-rows:auto auto !important; column-gap:12px !important; align-items:center !important;
  height:100% !important; box-sizing:border-box !important; }
body.${BODY_CLASS} .statusBtn .sapMBtnInner::before { content:attr(data-nd-ik); grid-column:1; grid-row:1 / span 2;
  width:40px; height:40px; border-radius:12px; background:rgba(255,255,255,.22);
  display:flex; align-items:center; justify-content:center; font-size:19px; line-height:1; }
body.${BODY_CLASS} .statusBtn .sapMBtnContent { grid-column:2 !important; grid-row:1 !important;
  justify-content:flex-start !important; text-align:left !important; font-size:15px !important;
  font-weight:800 !important; line-height:1.2 !important; white-space:nowrap !important; }
body.${BODY_CLASS} .statusBtn .sapMBtnInner::after { content:attr(data-nd-pod); grid-column:2; grid-row:2;
  text-align:left; font-size:12px; opacity:.85; line-height:1.25; white-space:nowrap; font-weight:600; }
body.${BODY_CLASS} .statusBtn .sapMBtnInner[data-nd-pod=""]::after { display:none; }

/* ---------- nadpisy sekcii ---------- */
.nd-nadpis { font:800 12px/1.3 -apple-system,"Segoe UI",Roboto,sans-serif; letter-spacing:.14em;
  text-transform:uppercase; color:#4a6285; margin:6px 0 8px 2px; }

/* ---------- pravy stlpec appky ---------- */
body.${BODY_CLASS} #WorkcenterDetail--Order_Status_Flexbox {
  background:transparent !important; border:0 !important; box-shadow:none !important;
  padding-top:2px !important; }

/* ---------- patka ---------- */
#${FOOTER_ID} { position:fixed; left:0; right:0; bottom:0; height:34px; z-index:4;
  display:flex; align-items:center; gap:10px; padding:0 18px; box-sizing:border-box;
  background:rgba(255,255,255,.82); backdrop-filter:blur(6px);
  border-top:1px solid #e3ebf5; font:12px/1 -apple-system,"Segoe UI",Roboto,sans-serif;
  color:#7d8ea8; }
#${FOOTER_ID} img { height:18px; display:block; }
#${FOOTER_ID} .nd-f-n { font-weight:800; color:#13315c; letter-spacing:.04em; }
#${FOOTER_ID} .nd-f-r { margin-left:auto; }
body.${BODY_CLASS} { padding-bottom:34px !important; box-sizing:border-box; }
`;
            document.head.appendChild(st);
        }

        /* --- horny pruh: logo a nazov vedla povodnych ovladacich prvkov --- */

        function hornyPruh() {
            if (document.getElementById(HEADER_ID)) return;
            const odhlasenie = document.querySelector('[id$="Button_Logout"]');
            const pruh = odhlasenie && (odhlasenie.closest('.sapMIBar') || odhlasenie.closest('.sapMBar') ||
                                        odhlasenie.closest('.sapMTB'));
            if (!pruh) return;

            const box = document.createElement('div');
            box.id = HEADER_ID;
            const img = document.createElement('img');
            img.src = LOGO_HF_MALE;
            img.alt = 'HF Slovakia';
            const txt = document.createElement('span');
            txt.className = 'nd-h-txt';
            const n = document.createElement('span'); n.className = 'nd-h-n'; n.textContent = 'HF SLOVAKIA';
            const s = document.createElement('span'); s.className = 'nd-h-s';
            s.textContent = 'Better Parts. A Cleaner Tomorrow.';
            txt.appendChild(n); txt.appendChild(s);
            box.appendChild(img); box.appendChild(txt);
            pruh.insertBefore(box, pruh.firstChild);
        }

        /* --- stavove tlacidla: ikona + nazov + podnadpis --- */

        /*
         * Len data-* atributy na obale .sapMBtnInner - ikonu a podnadpis z nich
         * vykresli CSS (::before / ::after). Do vnutra tlacidla sa nesiaha,
         * takze stlacenie funguje presne ako v povodnej appke.
         * Atributy nespustia DomWatch (sleduje len pridavanie/mazanie prvkov),
         * cyklus teda nehrozi; po prekresleni UI5 sa len znovu doplnia.
         */
        function kartaTlacidla(btn) {
            const inner = btn.querySelector('.sapMBtnInner');
            const content = btn.querySelector('.sapMBtnContent');
            if (!inner || !content) return;
            const text = (content.textContent || '').trim();
            if (!text) return;
            const k = kartaPre(text);
            if (inner.getAttribute('data-nd-ik') !== k.ikona) inner.setAttribute('data-nd-ik', k.ikona);
            if (inner.getAttribute('data-nd-pod') !== k.pod) inner.setAttribute('data-nd-pod', k.pod);
        }

        /* --- nadpisy sekcii --- */

        function nadpis(pred, text, znacka) {
            if (!pred || !pred.parentElement) return;
            const predch = pred.previousElementSibling;
            if (predch && predch.dataset && predch.dataset.ndNadpis === znacka) return;
            const d = document.createElement('div');
            d.className = 'nd-nadpis';
            d.dataset.ndNadpis = znacka;
            d.textContent = text;
            pred.parentElement.insertBefore(d, pred);
        }

        /* --- patka --- */

        function patka() {
            if (document.getElementById(FOOTER_ID)) return;
            const f = document.createElement('div');
            f.id = FOOTER_ID;
            const img = document.createElement('img');
            img.src = LOGO_HF_MALE;
            img.alt = '';
            const n = document.createElement('span'); n.className = 'nd-f-n'; n.textContent = 'HF SLOVAKIA';
            const s = document.createElement('span'); s.textContent = 'Better Parts. A Cleaner Tomorrow.';
            const r = document.createElement('span'); r.className = 'nd-f-r';
            let verzia = '';
            try { verzia = GM_info && GM_info.script ? 'v' + GM_info.script.version : ''; } catch (e) { /* ignore */ }
            r.textContent = ('PDA App Extension ' + verzia).trim() + ' · Powered by HF Slovakia';
            f.appendChild(img); f.appendChild(n); f.appendChild(s); f.appendChild(r);
            document.body.appendChild(f);
        }

        function apply() {
            injectStyles();
            if (!document.body.classList.contains(BODY_CLASS)) document.body.classList.add(BODY_CLASS);

            hornyPruh();
            patka();

            document.querySelectorAll('.statusBtn').forEach((b) => {
                try { kartaTlacidla(b); } catch (e) { /* kozmetika nesmie nic zhodit */ }
            });

            const stav = document.getElementById('WorkcenterDetail--Order_Status_Flexbox');
            if (stav) nadpis(stav, 'Stav operácie', 'stav');
            const hlavicka = document.getElementById('WorkcenterDetail--OrderHeader_FlexBox');
            if (hlavicka) nadpis(hlavicka, 'Zákazka a materiál', 'zakazka');
        }

        DomWatch.add(apply);
        onReady(apply);
    }

    /* -------------------- 3.19 Ladiaci vypis ---------------------------- */

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
            id: 'donut3d',
            name: 'Priestorové koláčové grafy',
            desc: 'Tri koláče s časmi SAP (Setup / Machine / Labor) nakloní ako pohľad zboku a pridá tieň. Dáta ani hodnoty sa nemenia, iba vzhľad.',
            def: true,
            run: modDonut3D,
        },
        {
            id: 'fullLeft',
            name: 'Ľavý panel na celú výšku',
            desc: '⚠️ Rozpracované — zatiaľ rozhadzuje rozloženie (stĺpec sa pripne na zlé miesto a panel pracoviska sa zosype). Nezapínať, kým to nedoladíme.',
            def: false,
            run: modFullHeightLayout,
        },
        {
            id: 'hfMenu',
            name: 'Menu HF Slovakia (vpravo)',
            desc: 'Zvislý panel pri pravom okraji s pripravovanými funkciami HF Slovakia (CHIPS, privolanie majstra, odvoz materiálu, TOOLSHOP, Flexus). Tlačidlá zatiaľ ukážu okno „vo vývoji“.',
            def: true,
            run: modHfMenu,
        },
        {
            id: 'newDesign',
            name: 'Nový dizajn HF Slovakia',
            desc: 'Prezlečenie celej aplikácie: svetlomodré pozadie, biele karty, horný pruh s logom, stavové tlačidlá s ikonou a podnadpisom, nadpisy sekcií, pätka. Všetky pôvodné ovládacie prvky ostávajú — menia sa len farby a tvary.',
            def: true,
            run: modNewDesign,
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
