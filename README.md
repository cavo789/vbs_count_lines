![Banner](images/banner.jpg)

# Count the number of lines

> Scan a folder, open every files based on a given extension, count the number of lines and display a summary

## Table of Contents

- [Description](#description)
- [Install](#install)
- [Usage](#usage)
- [Author](#author)
- [License](#license)

## Description

This small script aims one simple feature : open every .txt files inside the current folder and, for each files, open it, read it's content and count the number of lines.

The result of the script is simply to display in the command prompt each filenames and the number of lines.

Of course, this script can be used to desserve others needs like outputting infos into a log file, in a .csv file, ...

Just modify it to fit your need.

## Install

Just get a copy of the `count_lines.vbs` script, save it on a folder on your disk / network drive.

## Usage

From a DOS prompt, run the script with two parameters, for instance:

```
cscript.exe count_lines.vbs c:\temp txt
```

where `c:\temp\` is the folder to scan and `txt` the extension of files to scan.

## Output

Here is a sample of output

```markdown
| Filename             | Count |
| -------------------- | ----- |
| cbxLanguages.bas     | 112   |
| cbxProductOwners.bas | 188   |
| cbxSessionID.bas     | 271   |
| cbxSurveyID.bas      | 261   |
| Comments.bas         | 194   |
| Constants.bas        | 25    |
| edtEndDateFrom.bas   | 50    |
| edtEndDateTill.bas   | 50    |
| edtTitle.bas         | 48    |
| edtTrainer.bas       | 52    |
| edtTrainingCode.bas  | 53    |
| Graphs.bas           | 116   |
| Helpers.bas          | 298   |
| NoAnswers.bas        | 80    |
| Toolbar.bas          | 465   |
| Variables.bas        | 25    |
| WebHelpers.bas       | 3.177 |
| TOTAL                | 5.465 |
```

## Author

Christophe Avonture

## Contribute

PRs accepted.

## License

[MIT](LICENSE)
