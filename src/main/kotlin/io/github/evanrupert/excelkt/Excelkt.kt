package io.github.evanrupert.excelkt

import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFWorkbook

/**
 * Creates a new Excel workbook
 *
 * @param style optional cell style to be applied to all cells in the workbook
 * @param init block function where sheets can be added
 * @return the newly created XSSFWorkbook after the init function has been run
 */
fun workbook(style: XSSFCellStyle? = null, init: Workbook.() -> Unit): Workbook =
    Workbook(XSSFWorkbook(), style).apply(init)
