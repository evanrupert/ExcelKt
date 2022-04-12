package com.github.evanrupert.excelkt

import org.apache.poi.xssf.usermodel.*
import java.time.LocalDate
import java.time.LocalDateTime
import java.time.ZoneId
import java.util.*

@DslMarker
annotation class ExcelElementMarker

/**
 * Super class for all Excel elements
 * @property xssfWorkbook underlying apache poi workbook from which styles and fonts are created
 */
@ExcelElementMarker
abstract class ExcelElement {
    abstract val xssfWorkbook: XSSFWorkbook

    /**
     * Creates a new XSSFCellStyle from the options specified in the init block
     *
     * @param init block function where style attributes can be specified
     * @return newly created XSSFCellStyle
     */
    fun createCellStyle(init: XSSFCellStyle.() -> Unit = { }): XSSFCellStyle = xssfWorkbook.createCellStyle().apply(init)

    /**
     * Creates a new XSSFFont from the options specified in the init block
     *
     * @param init block function where font attributes can be specified
     * @return newly created XSSFFont
     */
    fun createFont(init: XSSFFont.() -> Unit = { }): XSSFFont = xssfWorkbook.createFont().apply(init)
}

/**
 * Represents an Excel workbook
 *
 * @property xssfWorkbook underlying apache poi workbook
 * @property style XSSFCellStyle applied to all lower elements
 */
class Workbook(
    override val xssfWorkbook: XSSFWorkbook,
    private val style: XSSFCellStyle?
) : ExcelElement() {
    /**
     * Creates a new sheet in the workbook
     *
     * @param name the name to be given to this sheet
     * @param style optional cell style to be applied to all lower elements
     * @param init block function where rows can be added
     */
    fun sheet(name: String? = null, style: XSSFCellStyle? = null, init: Sheet.() -> Unit) =
        Sheet(
            xssfWorkbook = xssfWorkbook,
            style = style ?: this.style,
            name = name
        ).apply(init)
}

/**
 * Represents a sheet in a workbook
 *
 * @property xssfWorkbook underlying apache poi workbook
 * @property style XSSFCellStyle applied to all lower elements
 * @property xssfSheet underlying apache poi xssfSheet
 * @param name name given to this sheet
 */
class Sheet(
    override val xssfWorkbook: XSSFWorkbook,
    private val style: XSSFCellStyle?,
    name: String?
) : ExcelElement() {
    val xssfSheet = name?.let(xssfWorkbook::createSheet) ?: xssfWorkbook.createSheet()
    private var currentRowIndex = 0

    /**
     * Creates a new row in the sheet
     *
     * @param style optional cell style to be applied to all cells in row
     * @param init block function where cells can be added
     */
    fun row(style: XSSFCellStyle? = null, init: Row.() -> Unit) =
        Row(
            xssfWorkbook = xssfWorkbook,
            style = style ?: this.style,
            xssfSheet = xssfSheet,
            index = currentRowIndex++
        ).apply(init)
}

/**
 * Represents a row in a sheet
 *
 * @property xssfWorkbook underlying apache poi workbook
 * @property style optional XSSFCellStyle applied to all cells in row
 * @param xssfSheet underlying apache poi xssfSheet where row was created
 * @param index the index on the sheet where row was created
 * @property xssfRow underlying apache poi xssfRow
 */
class Row(
    override val xssfWorkbook: XSSFWorkbook,
    private val style: XSSFCellStyle?,
    xssfSheet: XSSFSheet,
    index: Int
) : ExcelElement() {
    val xssfRow = xssfSheet.createRow(index)
    private var currentCellIndex = 0

    /**
     * Creates a new cell in the row
     *
     * Cell explicitly support the following data types:
     * - Formula
     * - Boolean
     * - Number
     * - Date
     * - Calendar
     * - LocalDate
     * - LocalDateTime
     *
     * All other data types will be converted to a string to be displayed.
     *
     * Example of all data types in use:
     * ```kotlin
     *
     * row {
     *   cell(Formula("A1 + A2"))
     *   cell(true)
     *   cell(12.2)
     *   cell(Date())
     *   cell(Calendar.getInstance())
     *   cell(LocalDate.now())
     *   cell(LocalDateTime.now())
     * }
     * ```
     *
     * @param content the content to be displayed in the cell
     * @param style optional cell style to be applied to this cell
     */
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

/**
 * Represents a cell in a row
 * @property xssfWorkbook underlying apache poi workbook
 * @property style XSSFCellStyle applied to this cell
 * @param content content displayed in this cell
 * @param xssfRow underlying apache poi row where this cell was created
 * @param index this cell's position in the row
 */
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

/**
 * Represents the formula data type for a cell
 *
 * Example:
 * ```kotlin
 * cell(Formula("C1 + B1"))
 * ```
 *
 * @property content string expression of an Excel formula
 */
data class Formula(val content: String)
