![File Renamer](./FileRenamer.png)

## Setup

Please read the following if you are unfamiliar with how to execute Powershell scripts on Windows.

[Run Powershell Scripts](https://adamtheautomator.com/run-powershell-script/)

[Launch Powershell as Administrator](https://adamtheautomator.com/powershell-run-as-administrator/)

[Install ImportExcel Module](https://github.com/dfinke/ImportExcel#installation)

## What It Does

Copies files into a new directory renaming them based on an Excel spreadsheet you provide.

## How To Use

```ps
cd C:\Users\YourName\Downloads\FileRenamer\Example
```

This above command depends on where you downloaded FileRenamer folder

## Example

Using defaults inside of directory with files

```ps
# FileRenamer/Example

..\RenameFiles.ps1
```

Same as above just showing CLI params for example purposes

```ps
# FileRenamer/Example

..\RenameFiles.ps1 -sheet 'Example.xlsx' -directory $pwd
```

Just run the same command again and it will delete your `Rename` directory with all your new file names.

> Don't worry your old files were still preserved.

## CLI

`[String] -sheet = 'Example.xlsx'`

Pass the path to the excel document that holds your old and new file names.

`[String] -directory = $pwd`

Pass the path where your old files are located.

> It is recommended to `cd` into the directory of interest where your files are located and then run the script from there.

`[String[]] -extensions = ('*.jpg', '*.jpeg', '*.png', '*.svg', '*.txt')`

Pass the extensions you would like to support.

## File Support

Theoretically you could support any standard type of file but for right now the following are supported

## Requirements

In order for this to work the following Excel document requirements are needed.

See the [Example Excel](./Example/Example.xlsx) document for reference.

  1. Excel spreadsheet can only have **1 workbook** with **2 columns** in A and B they need to be named **Old** and **New**
  2. In the Old Column (A) put the existing file names
  3. In the New Column (B) put the new file names

> Order matters for the columns, the names need to match (old vs new) and the length of the columns must be the same.

## Contribution

This is a super quick, simple example. Feel free to post issues, contribute, fork, whatever improves the quality of this lil script.