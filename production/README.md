# PDA Suite — produkčný build

Vylepšenia pre aplikáciu **PDA** bežiacu na `https://hf.simplifier.cloud/appDirect/PDA/`.

Skripty nemenia aplikáciu na serveri — bežia až v prehliadači a dopĺňajú do nej
UI: nový dizajn celej aplikácie, slovenské popisky, farebné stavové tlačidlá,
vyhľadávanie zákaziek naprieč pracoviskami, prehľad pracovísk na úvodnej
obrazovke, načítanie výkresov, obsluhu skenera a čítačky kariet.

---

## Inštalácia — stačí jeden skript

**[▶ Inštalovať PDA Suite](https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/production/pda-suite.user.js)**

`pda-suite.user.js` obsahuje **všetkých 18 modulov v jednom súbore**. Nainštaluješ
ho raz a jednotlivé moduly si potom zapínaš a vypínaš v nastaveniach — nemusíš
nič odinštalovávať ani doinštalovávať.

Súbor je **sebestačný**: pozadie aplikácie aj všetky obrázky na stavových
tlačidlách sú v ňom vložené ako `data:` URI, takže nepotrebuje žiadne ďalšie
súbory z repozitára. Jediná externá závislosť je knižnica `xlsx` z CDN
(`@require`), ktorú Tampermonkey stiahne sám a používa ju len modul výkresov.

### Predpoklady (raz za počítač)

1. **Tampermonkey** —
   [Chrome Web Store](https://chromewebstore.google.com/detail/tampermonkey/dhdgffkkebhmkfjojejmpbldmpobfkfo)
2. **Povoliť používateľské skripty** — `chrome://extensions` → Tampermonkey →
   *Podrobnosti* → zapnúť **„Allow user scripts"**.
   ⚠️ Bez tohto kroku sa skripty tvária, že bežia, ale nerobia nič.

V Tampermonkey nechaj **zapnutý vždy len jeden** skript PDA Suite — produkčný
build a staršie vetvy si navzájom prepisujú tie isté prvky v DOM.

### Overenie

Otvor PDA a daj **Ctrl + F5**. Vpravo dole sa objaví **⚙ ozubené koliesko** —
to je panel nastavení. V konzole prehliadača sa vypíše `[PDA Suite] aktívne
moduly: …`.

---

## Nastavenia

Panel otvoríš **ozubeným kolieskom vpravo dole**, alebo cez ikonu Tampermonkey →
*Nastavenia PDA Suite*. Pýta si heslo (predvolene `123456`, mení sa v paneli
v sekcii *Heslo do nastavení*).

| Modul | Predvolene | Poznámka |
|---|---|---|
| Nový dizajn HF Slovakia | zapnuté | prezlečenie celej aplikácie — pozadie, biele karty, nadpisy sekcií, stavové tlačidlá s obrázkom |
| Vylepšená hlavička | zapnuté | slovenské popisky pri ikonách, väčšie meno, zvýraznené odhlásenie |
| Farebné tlačidlá | zapnuté | farby podľa pravidiel v nastaveniach (výroba zelená, prestoj oranžový, chyba červená) |
| Vyhľadávanie zákaziek | zapnuté | hľadá naprieč všetkými pracoviskami naraz |
| Prehľad pracovísk na úvode | zapnuté | úvodná obrazovka zbalená do kategórií (Assembly, Welding, Machining, Quality Control) |
| Zoznam zákaziek ako pilulky | zapnuté | každá zákazka v jednom kompaktnom riadku, pevná výška zoznamu |
| Kompaktná hlavička detailu | zapnuté | zákazka, materiál a Production Order v jednom boxe; výkres, BOM a Operation Complete v jednom riadku |
| Popis operácie na celú obrazovku | zapnuté | veľké tlačidlo, po kliknutí popis cez celú obrazovku |
| Priestorové koláčové grafy | zapnuté | časy SAP (Setup / Machine / Labor) naklonené ako pohľad zboku |
| Krajší graf vyťaženia | zapnuté | čas dole len ako HH:MM, tenšie pásy, nižšie plátno |
| Krajšia tabuľka stavov | zapnuté | krátky dátum a čas, stav ako farebná pilulka, striedavo podfarbené riadky |
| Menu HF Slovakia (vpravo) | zapnuté | zvislý panel s pripravovanými funkciami (CHIPS, majster, materiál, TOOLSHOP, Flexus) |
| Blokovanie tlačidla Späť | zapnuté | zabráni nechcenému vypadnutiu z aplikácie |
| Panel prepínania používateľov | vypnuté | má zmysel len na zdieľanom termináli |
| Čiarový skener a RFID karty | vypnuté | vyžaduje hardvér |
| Tlačidlo výkresu | vypnuté | vyžaduje firemnú sieť a Excel so zoznamom výkresov |
| Ľavý panel na celú výšku | vypnuté | ⚠️ rozpracované, zatiaľ rozhadzuje rozloženie |
| Ladiaci výpis do konzoly | vypnuté | len na hľadanie chýb |

Zmena sa uloží tlačidlom **Uložiť a obnoviť stránku**. Nastavenia sa dajú
zálohovať a prenášať medzi terminálmi cez *Uložiť do súboru* / *Načítať zo
súboru*.

### Citlivé údaje

**Heslá kolegov, ID kariet ani kľúč k službe výkresov nie sú v kóde** a nikdy sem
nepatria — repozitár je verejný. Zadávajú sa v paneli nastavení a ukladajú sa len
lokálne do úložiska Tampermonkey na danom počítači.

V kóde ostávajú dve veci, ktoré skript potrebuje ako predvolené hodnoty:
predvolené heslo do panela nastavení (`123456`, po inštalácii si ho zmeň) a adresa
internej služby výkresov (`http://172.16.77.134:9000`), ktorá je zároveň
v hlavičke ako `@connect`.

---

## Ako to vnútri funguje

Celý skript stojí na troch spoločných častiach, ktoré zdieľajú všetky moduly:

- **`DomWatch`** — jeden `MutationObserver` pre celú stránku. Aplikácia je SAP
  UI5 a prekresľuje si DOM sama, takže moduly musia svoje prvky dokladať znova.
  Všetky prekreslenia sa zlučujú do jednej dávky.
- **`XhrBus`** — jedno odpočúvanie volaní na `/client/1.0/executeBO`. Prepisuje
  `XMLHttpRequest` v kontexte stránky a odpovede rozposiela modulom, ktoré sa
  prihlásili.
- **`UserSwitch`** — prepnutie používateľa cez SAP dialóg. Používa ho bočný
  panel aj čítačka kariet, preto je mimo oboch modulov.

Moduly medzi sebou komunikujú cez objekt `shared` (index zákaziek, aktuálna
operácia), nie cez globálne premenné na `window`.

Obrázky (pozadie aplikácie, grafika na stavových tlačidlách) sú konštanty
`POZADIE` a `OBR_*` hneď nad modulom nového dizajnu — sú to base64 `data:` URI,
takže build nezávisí na žiadnom súbore vedľa.

---

## Pridanie nového vylepšenia

1. Napísať funkciu `modNiecoNove()` v sekcii **3. MODULY**
2. Pridať jeden riadok do zoznamu `MODULES` (`id`, `name`, `desc`, `def`, `run`)
3. Zvýšiť `@version` v hlavičke — **inak Tampermonkey aktualizáciu nestiahne**
   (aktualizuje sa z `@updateURL`, teda z vetvy `main` tohto repozitára)
4. Commit a push

Modul sa automaticky objaví v paneli nastavení. Ak niektorý modul spadne,
ostatné bežia ďalej — spúšťajú sa každý vo vlastnom `try/catch`.

---

Pôvodný základ pochádza z repozitára `Dan1elG94/HF-Slovakia-PDA-scripts`.
