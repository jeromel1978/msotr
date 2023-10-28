# MSOTR

## This module is in very early development

## Description

Replace MS Office template placeholders with specified content

## Requirements

Placeholders in Office document must be encased in "{" and "}"

JSON used for replacing content do not need the curly braces
Special nodes:

- XLSX
  - used for PPTX embedded charts and related tables
- TABLES
  - Used for PPTX tables
  - Template Requires Header Row and two rows
  - Placeholder name must be somewhere in the table

Either URL or Local are required to import template

## Sample

JSON Sample format in "sample.json"

## Element Requirements

- XLSX Element requirements

  - Name - Placeholder name
  - Workbook - Workbook name for referencing embedded Excel file
  - Data - See sample for format

- TABLES Element requirements
  - Name - Placeholder name
  - Headers - Array of strings
  - Data - 2 dimentional array of string

# CLI Usage

node ./dist/msotr.js "[\path\to\input].json" "[\path\to\template].pptx" "[\path\to\output].pptx"

Function Parameters

- URL (Optional) - URL of template file
- Local (Optional) - local path of template file
- Replacements (Required) - JSON containing all placeholder names and values to replace with
- Out (Optional) - path to export filled template. If omitted, the function will return the filled template as Buffer
