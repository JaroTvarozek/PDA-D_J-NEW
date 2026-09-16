# NEW Design — druhá vetva PDA Suite

Tu sa vyvíja **nový dizajn**. Pôvodný skript v koreňovom priečinku ostáva nedotknutý a funkčný.

| | Pôvodná verzia | Nový dizajn |
|---|---|---|
| Súbor | `pda-suite.user.js` | `ver.2/new-design.user.js` |
| Názov v Tampermonkey | **PDA Suite (HF Slovakia)** | **PDA Suite NEW Design (HF Slovakia)** |
| Verzia | 1.25.x | 2.0.0 a vyššie |
| Výpis v konzole | `[PDA Suite]` | `[PDA NEW Design]` |
| Inštalácia / aktualizácia | [odkaz](https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/pda-suite.user.js) | [odkaz](https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/ver.2/new-design.user.js) |

Sú to pre Tampermonkey **dva samostatné skripty** (líšia sa `@name` aj `@namespace`), takže:

- aktualizujú sa nezávisle — zmena v jednom nemá na druhý žiadny vplyv,
- dajú sa zapínať a vypínať samostatne.

---

## Inštalácia nového dizajnu

1. Otvor [inštalačnú adresu](https://github.com/JaroTvarozek/PDA-D_J-NEW/raw/refs/heads/main/ver.2/new-design.user.js)
2. Tampermonkey ponúkne **Install** — potvrď
3. V Tampermonkey (ikona → **Dashboard**) budeš mať oba skripty v zozname

## Prepínanie medzi verziami

⚠️ **Vždy nechaj zapnutý iba jeden z nich.** Keby bežali oba naraz, obe by kreslili to isté
(dvakrát panel, dvakrát okienko výkresu) a obrazovka by bola zmätok.

1. Ikona Tampermonkey → **Dashboard**
2. V stĺpci **Enabled** prepni:
   - chcem pôvodný → *PDA Suite* **zapnutý**, *PDA Suite NEW Design* **vypnutý**
   - chcem nový → naopak
3. V PDA daj **Ctrl + F5**

Rýchlejšie: klikni na ikonu Tampermonkey priamo na stránke PDA — oba skripty sú v zozname
a prepínač je hneď pri názve.

---

## ⚠️ Nastavenia sa neprenášajú

Tampermonkey drží nastavenia **oddelene pre každý skript**. Nový dizajn teda začína
s prázdnymi nastaveniami: bez používateľov, s predvolenými farbami tlačidiel,
heslom `123456` a **bez adresy Excelu** (takže by nenašiel výkres).

Prenesieš ich takto:

1. V **pôvodnom** skripte: ⚙ vpravo dole → heslo → *Súbor s nastaveniami* → **Stiahnuť**
2. Prepni sa na **nový dizajn** (viď vyššie), obnov stránku
3. ⚙ → heslo → *Súbor s nastaveniami* → **Nahrať** → vyber stiahnutý `.json`

Alebo jednoducho znovu zadaj adresu Excelu v ⚙ → *Excel*.

---

## Ako sa tu pracuje

- Nový dizajn sa upravuje **iba v `ver.2/new-design.user.js`**
- Pôvodný `pda-suite.user.js` sa **nemení** — je to poistka, na ktorú sa dá kedykoľvek vrátiť
- Verzie sa číslujú od **2.0.0** nahor; `@version` treba zvýšiť pri každom vydaní,
  inak Tampermonkey aktualizáciu nestiahne
- Kontrola pred vydaním: `node --check ver.2/new-design.user.js`
  a `node .claude/nastroje/check.js ver.2/new-design.user.js`

Nový dizajn vznikol ako presná kópia verzie **1.25.4** pôvodného skriptu (2026-09-16),
takže na začiatku sa správa rovnako — odlišovať sa bude až tým, čo do neho pribudne.
