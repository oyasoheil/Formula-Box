
# Excel Formulas and short VBA codes

## Table of Contents
- [Introduction](#introduction)
- [VBA Codes](#vba-codes)
  - [Code 1: Function Name](#code-1-function-name)
  - [Code 2: Function Name](#code-2-function-name)
- [Formulas](#formulas)
  - [Formula 1: Description](#formula-1-description)
  - [Formula 2: Description](#formula-2-description)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Introduction
A brief overview of your VBA utility library, its purpose, and key features.

## VBA Codes

### Code 1: Function Name
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Count = 1 Then
        Application.EnableEvents = False
        Rows(Target.Row).Select
        Application.EnableEvents = True
    End If
End Sub

