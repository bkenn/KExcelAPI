package link.webarata3.kexcelapi

import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.xssf.usermodel.*
import java.awt.Color


fun Sheet.merge(firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int) {
    this.addMergedRegion(CellRangeAddress(firstRow, lastRow, firstCol, lastCol))
}

fun XSSFSheet.mergeAndStyle(firstRow: Int, lastRow: Int, firstCol: Int, lastCol: Int, op: XSSFCellStyle.() -> Unit) {
    this.addMergedRegion(CellRangeAddress(firstRow, lastRow, firstCol, lastCol))
    val style = workbook.createCellStyle()
    style.op()
    for(i in firstRow..lastRow) {
        this[i].forEach { it.cellStyle = style }
    }
}

fun XSSFSheet.style(op: XSSFCellStyle.() -> Unit) : XSSFCellStyle {
    val style = workbook.createCellStyle()
    style.op()
    return style
}
fun XSSFWorkbook.style(op: XSSFCellStyle.() -> Unit) : XSSFCellStyle {
    val style = createCellStyle()
    style.op()
    return style
}

fun XSSFSheet.font(op: Font.() -> Unit): Font {
    val font = workbook.createFont()
    font.op()
    return font
}

fun XSSFWorkbook.font(op: XSSFFont.() -> Unit): XSSFFont {
    val font = createFont()
    font.op()
    return font
}

fun XSSFWorkbook.datahelper(format: String): Short {
    return creationHelper.createDataFormat().getFormat(format)
}

fun XSSFSheet.styleRow(row: Int, style: XSSFCellStyle) {
    this[row].forEach { it.cellStyle = style }
}

fun XSSFCellStyle.setFillForegroundColor(color: Color) = setFillForegroundColor(XSSFColor(color))
