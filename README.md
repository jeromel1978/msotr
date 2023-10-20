# MSOTR

## This module is in very early development

## Description

Replace MS Office template placeholders with specified content

## Requirements

Placeholders in Office document must be encased in "{" and "}"

JSON used for replacing content do not need the curly braces
Special nodes:

- XLSX - used for PPTX embedded charts and related tables
- TABLES - Used for PPTX tables

Either URL or Local are required to import template

# CLI Usage

node ./dist/msotr.js "[\path\to\input].json" "[\path\to\template].pptx" "[\path\to\output].pptx"

Function Parameters

- URL (Optional) - URL of template file
- Local (Optional) - local path of template file
- Replacements (Required) - JSON containing all placeholder names and values to replace with
- Out (Optional) - path to export filled template. If omitted, the function will return the filled template as Buffer
