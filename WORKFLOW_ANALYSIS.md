# MotoCalc Workflow Analysis

This document explains the actual data flow in the provided Office files. It is based on the workbook contents, extracted VBA, and the current Python replacement app.

## 1. Short Answer

The real source of data is [6055.xls](6055.xls). It is the master form where the user edits the vehicle and person data.

The other files are targets/templates:

- [6055_MOTO_template.xls](6055_MOTO_template.xls) is the ACT template.
- [DOGOVIR_6055_template.doc](DOGOVIR_6055_template.doc) is the contract template.
- [vidatkova.xls](vidatkova.xls) is a separate invoice/expense template and is not part of the 6055 -> act -> contract chain.

So the direction is not “templates into 6055”. It is:

1. User fills [6055.xls](6055.xls).
2. The act is generated from [6055_MOTO_template.xls](6055_MOTO_template.xls).
3. The contract is generated from [DOGOVIR_6055_template.doc](DOGOVIR_6055_template.doc).
4. [vidatkova.xls](vidatkova.xls) is independent.

## 2. File-by-File Structure

### 2.1 [6055.xls](6055.xls)

Single sheet: `Worksheet`

Role: master source workbook. No VBA or XLM macros were found.

Key fields:

- `A3`: document number + date line, for example `№ 8424/26/009999 від 15 травня 2026 року`
- `C12`: birthday
- `C15`: full buyer name
- `C16`: address
- `C17`: passport info
- `C18`: tax number
- `C20`: vehicle model
- `C21`: vehicle type
- `C22`: vehicle purpose
- `C25`: category
- `C26`: body type
- `C27`: fuel type
- `C28`: color
- `C29`: year
- `C30`: gross weight
- `C31`: curb weight
- `C33`: seats
- `C35`: cylinders
- `C36`: engine capacity
- `C37`: engine power in kW
- `C39`: VIN
- `C41`: engine number
- `C42`: chassis number
- `C43`: frame number
- `C44`: price without VAT
- `C45`: VAT
- `C46`: price with VAT
- `C47`: customs declaration
- `C50`: plate/sign
- `C51`: acquisition date
- `C53`: conclusion text
- `A56` / `D56`: seller/buyer signature names

Important detail: `C39`, `C42`, and `C43` all contain the same frame/VIN-like value in the current file. The workflow treats these as synchronized identity fields.

### 2.2 [6055_MOTO_template.xls](6055_MOTO_template.xls)

Single sheet: `Worksheet`

Role: act template. This file contains the meaningful VBA logic.

It has these VBA parts:

- `ЭтаКнига.cls`: empty
- `Лист1.cls`: empty
- `UserForm1.frm`: empty button handler
- `Module1.bas`: the real logic

#### Visible sheet structure

This sheet is a formatted act document, not a raw data table. The macro fills specific cells in the form layout:

- `B12`: number parsed from `6055.xls!A3`
- `L12`: date parsed from `6055.xls!A3`
- `H18`: full name from `6055.xls!C15`
- `C24`: model from `6055.xls!C20`
- `E24`: year from `6055.xls!C29`
- `H24`: chassis/VIN from `6055.xls!C39`
- `K24`: color from `6055.xls!C28`
- `N33`: shortened name derived from `6055.xls!C15`

#### Macro inventory

##### `СУММАПРОПИСЬЮ(n As Double) As String`

Helper function that converts a numeric amount into Ukrainian words.

What it does:

- Splits the number into digits using `Class`.
- Builds words for millions, thousands, hundreds, tens, and ones.
- Returns the textual representation.

Important note:

- In this workbook version, it returns only the words.
- The Word contract workflow later appends the currency text in the Python replacement, while the original VBA contract flow writes this result into the `sumtext` bookmark.

##### `Class(M, i)`

Private helper used by `СУММАПРОПИСЬЮ`.

It extracts the digit at position `i` from a number.

##### `MyCopy(Из As String, В As String)`

Copies a range value from `6055.xls` into a range in `6055_MOTO_template.xls`.

Reads from:

- `Workbooks("6055.xls").Worksheets("Worksheet").Range(Из)`

Writes to:

- `Workbooks("6055_MOTO_template.xls").Worksheets("Worksheet").Range(В)`

This is a direct cell-to-cell transfer helper.

##### `MyCopyNumberAndDate(Fr As String, T1 As String, T2 As String)`

Reads `A3` from `6055.xls`, splits the string by spaces, and extracts:

- `arr(1)` into `T1`
- `arr(3) & " " & arr(4) & " " & arr(5) & " " & arr(6)` into `T2`

In practice this means:

- number -> `B12`
- date text -> `L12`

##### `MyCopyIn(Из, В)`

Copies a value from `6055.xls` to the current workbook and then shortens the full name into initials.

Behavior:

- Copies source text into the destination cell.
- Finds the first and last spaces.
- Rewrites the cell as `Surname N. P.` style initials.

This helper is used for `E15` and then the result is copied to `N33`.

##### `MyCopyInWord(w, Из, В)`

Takes a value from `6055.xls` and writes it into a Word bookmark.

The target is a bookmark name, not a cell.

##### `MyCopyInWordFromVar(w, Var, B)`

Writes a passed variable value into a Word bookmark.

Used for the parsed number and date from `A3`.

##### `GetNumber(Fr As String) As String`

Reads `A3`, splits it, returns `arr(1)`.

Used to place the number into the contract bookmark `Number`.

##### `GetDate(Fr As String) As String`

Reads `A3`, splits it, returns `arr(3) & " " & arr(4) & " " & arr(5) & " " & arr(6)`.

Used to place the date into the contract bookmark `Data`.

##### `Кнопка1_Щелчок()`

This is the main ACT-generation macro.

Read/write flow:

1. Opens `6055.xls` from the same folder.
2. Fills the act template in `6055_MOTO_template.xls`.
3. Closes `6055.xls`.

Actual transfers performed:

- `A3` -> `B12:C12` and `L12:N12` via `MyCopyNumberAndDate`
- `C15:D15` -> `H18:N18`
- `C20:D20` -> `C24:D25`
- `C29:D29` -> `E24:G25`
- `C39:D39` -> `H24:J25`
- `C28:D28` -> `K24:N25`
- `C15` -> `E15` -> `N33`

Formatting applied to `N33`:

- font size 9
- font color yellow
- font name Times New Roman
- bold
- borders removed

##### `Кнопка2_Щелчок()`

This is the Word contract-generation macro.

Read/write flow:

1. Opens `6055.xls`.
2. Creates `Word.Application` via `CreateObject`.
3. Opens `DOGOVIR_6055_template.doc`.
4. Writes values into Word bookmarks.
5. Leaves the Word document open in the original VBA version.
6. Closes `6055.xls`.

Bookmarks written:

- `Number` from parsed `A3`
- `Data` from parsed `A3`
- `FIO` from `C15`
- `pasport` from `C17`
- `TaxNumber` from `C18`
- `BirthDay` from `C12`
- `adres` from `C16`
- `decl` from `C47`
- `model` from `C20`
- `year` from `C29`
- `color` from `C28`
- `numberdv` from `C41`
- `cuzov` from `C39`
- `cub` from `C36`
- `znak` from `C50`
- `price` from `C46`
- `sumtext` from `СУММАПРОПИСЬЮ(C46)`
- `FIO2` from `C15`

Important nuance:

- This macro is the real reason `6055_MOTO_template.xls` depends on the contract template.
- Without `DOGOVIR_6055_template.doc`, this button cannot complete the document generation workflow.

### 2.3 [vidatkova.xls](vidatkova.xls)

Sheets:

- `Лист1`
- `Лист2`
- `Лист3`

Role: a separate invoice/expense-style template.

What it contains:

- Its own copy of `СУММАПРОПИСЬЮ`.
- No visible data-transfer macros.
- No references to `6055.xls`, `6055_MOTO_template.xls`, or `DOGOVIR_6055_template.doc`.

Visible sheet content shows a filled invoice-like form with supplier, recipient, order, line item table, totals, and signoff fields.

Important detail:

- The “formula-looking” strings like `${formatter.roundToCents(...)}` and `$[K10 - J10]` are literal cell text in the extracted file, not active Excel formulas in the current snapshot.
- The sheet is therefore best understood as a template or export artifact, not as an active macro-driven dependency of the 6055 workflow.

Conclusion:

- `vidatkova.xls` is standalone.
- It does not depend on the 6055/MOTO/contract chain.

### 2.4 [DOGOVIR_6055_template.doc](DOGOVIR_6055_template.doc)

Role: Word template with bookmarks.

It is not a macro workbook. The workflow uses bookmarks as insertion points.

Confirmed bookmark targets from the original VBA and the Python replacement:

- `Number`
- `Data`
- `FIO`
- `pasport`
- `TaxNumber`
- `BirthDay`
- `adres`
- `decl`
- `model`
- `year`
- `color`
- `numberdv`
- `cuzov`
- `cub`
- `znak`
- `price`
- `sumtext`
- `FIO2`

The document content is filled by replacing bookmark ranges with text.

## 3. Full Data Map

### From [6055.xls](6055.xls) to [6055_MOTO_template.xls](6055_MOTO_template.xls)

| Source | Target | Meaning |
|---|---|---|
| `A3` | `B12` | number part of document ID |
| `A3` | `L12` | date part of document ID |
| `C15` | `H18` | full buyer name |
| `C20` | `C24` | vehicle model |
| `C29` | `E24` | year |
| `C39` | `H24` | chassis / frame / VIN value |
| `C28` | `K24` | color |
| `C15` | `E15` | source for short name calculation |
| `E15` | `N33` | shortened initials version |

### From [6055.xls](6055.xls) to [DOGOVIR_6055_template.doc](DOGOVIR_6055_template.doc)

| Source | Bookmark | Meaning |
|---|---|---|
| `A3` parsed | `Number` | document number |
| `A3` parsed | `Data` | date |
| `C15` | `FIO` | full buyer name |
| `C17` | `pasport` | passport |
| `C18` | `TaxNumber` | tax ID |
| `C12` | `BirthDay` | birthday |
| `C16` | `adres` | address |
| `C47` | `decl` | declaration |
| `C20` | `model` | model |
| `C29` | `year` | year |
| `C28` | `color` | color |
| `C41` | `numberdv` | engine number |
| `C39` | `cuzov` | chassis/frame |
| `C36` | `cub` | cubic capacity |
| `C50` | `znak` | plate/sign |
| `C46` | `price` | numeric price |
| `C46` through `СУММАПРОПИСЬЮ` | `sumtext` | price in words |
| `C15` | `FIO2` | duplicate full name |

### Why `vidatkova.xls` is separate

`vidatkova.xls` does not appear in the macro call chain of `6055_MOTO_template.xls`.

The workflow in the existing VBA code never opens it, never writes to it, and never reads from it.

That means it is either:

- a separate business template,
- a historical file copied from another process,
- or a parallel invoice workflow that is unrelated to the 6055 act/contract generation.

## 4. Macro Inventory

### Real transfer macros

- `Кнопка1_Щелчок` in [6055_MOTO_template.xls](6055_MOTO_template.xls)
- `Кнопка2_Щелчок` in [6055_MOTO_template.xls](6055_MOTO_template.xls)

### Helper macros/functions in [6055_MOTO_template.xls](6055_MOTO_template.xls)

- `СУММАПРОПИСЬЮ`
- `Class`
- `MyCopy`
- `MyCopyNumberAndDate`
- `MyCopyIn`
- `MyCopyInWord`
- `MyCopyInWordFromVar`
- `GetNumber`
- `GetDate`

### Helper macros/functions in [vidatkova.xls](vidatkova.xls)

- `СУММАПРОПИСЬЮ`
- `Class`

### Empty or vestigial items

- `ЭтаКнига.cls` in both Excel files: empty
- sheet class modules in both Excel files: empty
- `UserForm1.frm` button click in `6055_MOTO_template.xls`: empty

## 5. Hidden / XLM Macro Check

What was found:

- `6055.xls`: no VBA, no XLM macros
- `6055_MOTO_template.xls`: VBA only, no working XLM workflow found
- `vidatkova.xls`: VBA helper only, no working XLM workflow found

What the scanner flagged:

- `Open`
- `Call`
- `CreateObject`
- hex/base64 patterns

These are normal Office automation patterns in this case, not proof of malicious behavior.

The `xlm_macro.txt` label in the `olevba` output for `vidatkova.xls` is a scanner artifact / metadata stream report, not evidence that the current business workflow relies on Excel 4 macros.

## 6. What Actually Has to Exist for the Workflow to Work

Minimum required files:

1. [6055.xls](6055.xls)
2. [6055_MOTO_template.xls](6055_MOTO_template.xls)
3. [DOGOVIR_6055_template.doc](DOGOVIR_6055_template.doc)

Optional / unrelated:

- [vidatkova.xls](vidatkova.xls)

## 7. What the Current Python App Already Does

The current app in [app_6055.py](app_6055.py) already reproduces the practical data flow:

- reads values from `6055.xls`
- generates the act output from `6055_MOTO_template.xls`
- fills the contract bookmarks on Windows
- creates a preview text on Linux

This means the app can be turned into a one-window UI without needing the original VBA buttons.

## 8. Recommendation for a Single-Window App

The cleanest structure is:

- One form with all editable fields from `6055.xls`
- One button to save source data
- One button to generate the act
- One button to generate the contract
- One button to generate all outputs
- One status/log panel

The app should treat `6055.xls` as the editable master record and the other documents as export targets, not as peer-editable files.

That is the main architectural simplification:

- source data in one place
- generation logic in one place
- outputs as derived artifacts
