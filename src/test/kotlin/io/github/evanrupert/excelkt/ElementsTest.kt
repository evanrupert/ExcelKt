package io.github.evanrupert.excelkt

import org.mockito.kotlin.*
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.xssf.usermodel.*
import org.junit.Test
import strikt.api.expectThat
import strikt.assertions.isEqualTo
import java.time.LocalDate
import java.time.ZoneId
import java.util.*
import org.mockito.kotlin.check as argCheck

class ElementsTest {
    private val mockXSSFFont: XSSFFont = mock()
    private val mockXSSFCellStyle: XSSFCellStyle = mock()
    private val mockXSSFCell: XSSFCell = mock()
    private val mockXSSFRow: XSSFRow = mock {
        on { createCell(org.mockito.kotlin.any()) } doReturn mockXSSFCell
    }
    private val mockXSSFSheet: XSSFSheet = mock {
        on { createRow(org.mockito.kotlin.any()) } doReturn mockXSSFRow
    }
    private val mockXSSFWorkbook: XSSFWorkbook = mock {
        on { createSheet() } doReturn mockXSSFSheet
        on { createFont() } doReturn mockXSSFFont
        on { createCellStyle() } doReturn mockXSSFCellStyle
    }

    private val wb: Workbook = Workbook(mockXSSFWorkbook, null)

    @Test
    fun `the sheet method creates a new sheet`() {
        wb.apply {
            sheet {}
        }

        verify(mockXSSFWorkbook).createSheet()
    }

    @Test
    fun `the workbook can create a new cell style`() {
        wb.createCellStyle()

        verify(mockXSSFWorkbook).createCellStyle()
    }

    @Test
    fun `the workbook can create a new font`() {
        wb.createFont()

        verify(mockXSSFWorkbook).createFont()
    }

    @Test
    fun `the sheet method can optionally take in a name`() {
        wb.apply {
            sheet("Test Sheet") {}
        }

        verify(mockXSSFWorkbook).createSheet("Test Sheet")
    }

    @Test
    fun `the row method can create multiple rows with proper indices`() {
        wb.apply {
            sheet {
                row {}
                row {}
                row {}
            }
        }

        verify(mockXSSFSheet).createRow(eq(0))
        verify(mockXSSFSheet).createRow(eq(1))
        verify(mockXSSFSheet).createRow(eq(2))
    }

    @Test
    fun `the cell method can create multiple cells with proper indices and proper content`() {
        wb.apply {
            sheet {
                row {
                    cell("Hello, First Cell!")
                    cell("Hello, Second Cell!")
                    cell(100.0)
                    cell(LocalDate.of(2021, 4, 28))
                    cell(Formula("A1+A2"))
                }
            }
        }

        verify(mockXSSFRow).createCell(eq(0))
        verify(mockXSSFRow).createCell(eq(1))
        verify(mockXSSFRow).createCell(eq(2))
        verify(mockXSSFRow).createCell(eq(3))
        verify(mockXSSFRow).createCell(eq(4))

        verify(mockXSSFCell).setCellValue(eq("Hello, First Cell!"))
        verify(mockXSSFCell).setCellValue(eq("Hello, Second Cell!"))
        verify(mockXSSFCell).setCellValue(eq(100.0))
        verify(mockXSSFCell).setCellValue(
            eq(
                Date.from(
                    LocalDate.of(2021, 4, 28).atStartOfDay(ZoneId.systemDefault()).toInstant()
                )
            )
        )
        verify(mockXSSFCell).setCellFormula(eq("A1+A2"))
    }

    @Test
    fun `the cell style is properly being set`() {
        val style = createStyleWith { fillBackgroundColor = IndexedColors.AQUA.index }

        wb.apply {
            sheet {
                row {
                    cell("Hello, World!", style)
                }
            }
        }

        verify(mockXSSFCell).cellStyle = argCheck {
            expectThat(it.fillBackgroundColor).isEqualTo(IndexedColors.AQUA.index)
        }
    }

    @Test
    fun `the cell style should be able to be set via a passed in lambda`() {
        wb.apply {
            val style = createCellStyle { fillBackgroundColor = IndexedColors.AQUA.index }

            sheet {
                row {
                    cell("Hello, World!", style)
                }
            }
        }

        verify(mockXSSFCell).cellStyle = any()
        verify(mockXSSFCellStyle).fillBackgroundColor = IndexedColors.AQUA.index
    }

    @Test
    fun `a font should be able to be created via a passed in lambda`() {
        wb.apply {
            val font = createFont { color = IndexedColors.AQUA.index }
            val style = createCellStyle { setFont(font) }

            sheet {
                row {
                    cell("Hello, World!", style)
                }
            }
        }

        verify(mockXSSFFont).color = IndexedColors.AQUA.index
        verify(mockXSSFCellStyle).setFont(any())
    }

    @Test
    fun `styles are passed down from the top if not overrode`() {
        val style = createStyleWith { fillBackgroundColor = IndexedColors.AQUA.index }

        val stylizedWorkbook = Workbook(mockXSSFWorkbook, style)

        stylizedWorkbook.apply {
            sheet {
                row {
                    cell("Hello, World!")
                }
            }
        }

        verify(mockXSSFCell).cellStyle = argCheck {
            expectThat(it.fillBackgroundColor).isEqualTo(IndexedColors.AQUA.index)
        }
    }

    @Test
    fun `lower element styles overwrite higher element styles`() {
        val style = createStyleWith { fillBackgroundColor = IndexedColors.AQUA.index }
        val newStyle = createStyleWith { fillBackgroundColor = IndexedColors.TEAL.index }

        wb.apply {
            sheet(style = style) {
                row {
                    cell("Hello, World!", newStyle)
                }
            }
        }

        verify(mockXSSFCell).cellStyle = argCheck {
            expectThat(it.fillBackgroundColor).isEqualTo(IndexedColors.TEAL.index)
        }

    }

    private fun createStyleWith(f: XSSFCellStyle.() -> Unit): XSSFCellStyle =
        XSSFWorkbook().createCellStyle().apply(f)
}
