package excelkt

import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileOutputStream

fun workbook(style: XSSFCellStyle? = null, init: Workbook.() -> Unit): XSSFWorkbook =
    Workbook(XSSFWorkbook(), style).apply(init).xssfWorkbook

fun XSSFWorkbook.write(filename: String) {
    val out = FileOutputStream(filename)
    write(out)
    out.close()
}
