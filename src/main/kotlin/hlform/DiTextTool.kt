package hlform

import com.itextpdf.text.*
import com.itextpdf.text.pdf.BaseFont
import com.itextpdf.text.pdf.PdfPCell


/**
 * 常用组件及工具
 *
 * @author
 * @date 2018-02-02
 */
object DiTextTool {

    private const val LOGO_PATH = "src/main/resources/logo.png"

    private val chineseFont: BaseFont
        get() = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED)

    private val titleFont = Font(chineseFont, 14f, Font.BOLD)
    private val contentFont = Font(chineseFont, 11f, Font.NORMAL)
    private val noteFont = Font(chineseFont, 10f, Font.NORMAL, BaseColor.GRAY)
    private val enFont = Font(Font.FontFamily.TIMES_ROMAN, 11f, Font.UNDERLINE)

    /**
     * 创建标题
     * @param document 文档
     */
    fun createTitle(document: Document) {
        val title = Paragraph("报价审批表", titleFont).apply {
            alignment = Element.ALIGN_CENTER
            //spacingBefore = 1f
        }
        document.add(title)
    }

    /**
     * 创建表头编号
     * @param document 文档
     * @param number 表单实体
     */
    fun createManageNumber(document: Document, number: String) {
        val paragraph = Paragraph("管理编号:  ", contentFont).apply {
            add(Phrase(number, enFont))
            alignment = Element.ALIGN_RIGHT
        }
        document.add(paragraph)
    }

    /**
     * 创建表头LOGO
     * @param document 文档
     */
    fun createLogo(document: Document) {
        val image = Image.getInstance(LOGO_PATH).apply {
            scaleToFit(100f, 200f)  //大小
            setAbsolutePosition(document.left(), document.top()) //左上角
            alignment = Element.ALIGN_LEFT
        }
        document.add(image)
    }

    /**
     * 表单表格部分单元格
     * @param value 单元格内容
     * @param colSpan 单元格行宽
     * @param rowSpan 单元格列宽
     * @param height 单元格高度
     * @return
     */
    fun cell(value: String? = null, colSpan: Int? = null, rowSpan: Int? = null, height: Float? = null): PdfPCell {
        return PdfPCell(Paragraph(value ?: "", contentFont)).apply {
            paddingTop = 3f
            paddingBottom = 3f
            horizontalAlignment = Element.ALIGN_LEFT
            verticalAlignment = Element.ALIGN_MIDDLE
            colSpan?.let { colspan = it }
            rowSpan?.let { rowspan = it }
            height?.let { fixedHeight = it }
        }
    }

    /**
     * 表单备注部分单元格
     * @param value 单元格内容
     * @param colSpan 单元格行宽
     * @param rowSpan 单元格列宽
     * @return
     */
    fun noteCell(value: String? = null, colSpan: Int? = null, rowSpan: Int? = null): PdfPCell {
        return PdfPCell(Paragraph(value ?: "", noteFont)).apply {
            verticalAlignment = Element.ALIGN_TOP
            border = 0
            colSpan?.let { colspan = it }
            rowSpan?.let { rowspan = it }
        }
    }

}