# BI_SDG9

This project contains ETL scripts for the SDG 9 datasets covering ICT usage,
telecommunications, and patents.

The Excel extraction layer does not rely on `openpyxl`. Each `.xlsx` file is
read with `zipfile.ZipFile`, treating the workbook as a compressed Open XML
package. The scripts then parse workbook XML parts such as
`xl/workbook.xml`, `xl/sharedStrings.xml`, and worksheet XML files with
`xml.etree.ElementTree`.
