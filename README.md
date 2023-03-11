# Excel-Utilities

This is a collection of personal utility functions for Excel.

## How to Use

1. Download the `.bas` with the functions that are needed.
2. Open a blank workbook and import the file in Excel VBA as a module.
3. Save the workbook as an add-in `.xlam` wherever you store add-ins for Excel (default location: `C:\Users\[username]\AppData\Roaming\Microsoft\AddIns`)
4. In Excel, go to `Options -> Add-ins -> Manage` and then enable the add-in.

## Public Functions

### Get SQL String

`getSQLString(inputText As String, Optional includeCommas As Boolean = True) As String` -> returns a string formatted to fit into a SQL query. Specifically, it escapes quotes, adds single quotes around the string, and optionally includes a comma at the end should you have multiple values that need copying (e.g., for a `VALUES` clause).

### Get UUID

`getUUID() As String` -> returns a psuedo-random, valid [version 4 UUID](https://en.wikipedia.org/wiki/Universally_unique_identifier).

### Get Workbook Name

`getWorkbookName() As String` -> returns the full name of the current workbook.

### To JSON

`toJSON(ParamArray vals() As Variant) As String` -> requires you input alternating name and value fields and returns a JSON string. E.g.,

||A|B|C|
|---|---|---|---|
|1|Fruit|Color|JSON|
|2|Apple|Red|=toJSON("Fruit",A2,"Color",B2)

**Result**

||A|B|C|
|---|---|---|---|
|1|Fruit|Color|JSON|
|2|Apple|Red|{"Fruit":"Apple","Color":"Red"}

### To JSON With Headers

`toJSONWithHeaders(selectedCells As Range) As String` -> requires a selection of cells and will use whatever values are in row 1 as the names for the values selected. E.g.,

||A|B|C|
|---|---|---|---|
|1|Fruit|Color|JSON|
|2|Apple|Red|=toJSONWithHeaders(A2:B2)

**Result**

||A|B|C|
|---|---|---|---|
|1|Fruit|Color|JSON|
|2|Apple|Red|{"Fruit":"Apple","Color":"Red"}



## Private Functions

### Get UUID Binary

`getUUIDBinary() As String` is an internal function that supports `getUUID()`. This function iterates over 128 bits to create a binary string to be used as the UUID. It has special case functions that ensure certain bit flags are set to certain values ensuring the generated UUID is a valid version 4 UUID.