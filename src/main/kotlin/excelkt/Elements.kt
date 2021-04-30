package excelkt

import org.apache.poi.xssf.usermodel.*
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.util.*

@DslMarker
annotation class ExcelElementMarker

@ExcelElementMarker
abstract class ExcelElement {
    abstract val xssfWorkbook: XSSFWorkbook
    fun createCellStyle(f: XSSFCellStyle.() -> Unit = { }): XSSFCellStyle = xssfWorkbook.createCellStyle().apply(f)
    fun createFont(f: XSSFFont.() -> Unit = { }): XSSFFont = xssfWorkbook.createFont().apply(f)
}

class Workbook(
    override val xssfWorkbook: XSSFWorkbook,
    private val style: XSSFCellStyle?
) : ExcelElement() {
    fun sheet(name: String? = null, style: XSSFCellStyle? = null, init: Sheet.() -> Unit) =
        Sheet(
            xssfWorkbook = xssfWorkbook,
            style = style ?: this.style,
            name = name
        ).apply(init)
}

class Sheet(
    override val xssfWorkbook: XSSFWorkbook,
    private val style: XSSFCellStyle?,
    name: String?
) : ExcelElement() {
    val xssfSheet = name?.let(xssfWorkbook::createSheet) ?: xssfWorkbook.createSheet()
    private var currentRowIndex = 0

    fun row(style: XSSFCellStyle? = null, init: Row.() -> Unit) =
        Row(
            xssfWorkbook = xssfWorkbook,
            style = style ?: this.style,
            xssfSheet = xssfSheet,
            index = currentRowIndex++
        ).apply(init)
}

class Row(
    override val xssfWorkbook: XSSFWorkbook,
    private val style: XSSFCellStyle?,
    xssfSheet: XSSFSheet,
    index: Int
) : ExcelElement() {
    val xssfRow = xssfSheet.createRow(index)
    private var currentCellIndex = 0

    fun cell(content: Any, style: XSSFCellStyle? = null) {
        Cell(
            xssfWorkbook = xssfWorkbook,
            style = style ?: this.style,
            content = content,
            xssfRow = xssfRow,
            index = currentCellIndex++
        )
    }
}

class Cell(
    override val xssfWorkbook: XSSFWorkbook,
    private val style: XSSFCellStyle?,
    content: Any,
    xssfRow: XSSFRow,
    index: Int
) : ExcelElement() {
    init {
        xssfRow.createCell(index).run {

            when (content) {
                is Formula -> setCellFormula(content.content)
                is Boolean -> setCellValue(content)
                is Number -> setCellValue(content.toDouble())
                is Date -> setCellValue(content)
                is Calendar -> setCellValue(content)
                is LocalDate -> setCellValue(Date.from(content.atStartOfDay(ZoneId.systemDefault()).toInstant()))
                is LocalDateTime -> setCellValue(Date.from(content.atZone(ZoneId.systemDefault()).toInstant()))
                else -> setCellValue(content.toString())
            }

            this@Cell.style?.let {
                cellStyle = it
            }
        }
    }
}

data class Formula(val content: String)
