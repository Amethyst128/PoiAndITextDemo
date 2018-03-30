package form

import com.itextpdf.text.Document
import com.itextpdf.text.DocumentException
import com.itextpdf.text.Element
import com.itextpdf.text.pdf.PdfPCell
import com.itextpdf.text.pdf.PdfPTable
import com.itextpdf.text.pdf.PdfWriter
import form.DiTextTool.cell
import form.DiTextTool.noteCell
import java.io.FileOutputStream
import java.io.IOException


/**
 * Description
 *
 * @author
 * @date 2018-02-02
 */
class QuoteApprovalForm {

    private val _file = "src/files/quoteForm.pdf"

    //生成表单;
    //表单中的label可使用国际化替换;方便不同语言表单之间切换;
    fun createForm() {
        try {
            val document = Document()
            PdfWriter.getInstance(document, FileOutputStream(_file))
            document.open()

            DiTextTool.createLogo(document)
            DiTextTool.createTitle(document)
            DiTextTool.createManageNumber(document, "SIC-FM032")
            this.createTable(document)
            this.addNotes(document)

            document.close()
        } catch (e: DocumentException) {
            e.printStackTrace()
        } catch (e: IOException) {
            e.printStackTrace()
        }
    }

    //创建表格
    private fun createTable(document: Document) {
        // val widths = floatArrayOf(10f, 25f, 30f, 30f)  // 设置表格的列宽和列数 默认是4列
        // 建立一个pdf表格
        val table = PdfPTable(10)
        table.spacingBefore = 10f
        table.widthPercentage = 100f

        //一个column表示一行
        var column = arrayOf(
            cell("申请人"),
            cell("张三"),
            cell("所属部门"),
            cell("研发部"),
            cell("岗位"),
            cell("打杂"),
            cell("申请日期"),
            cell(),
            cell("客户名称"),
            cell()
        )
        table.addAllCells(column)

        column = arrayOf(
            cell("产品名称"), cell(), cell("客户预计年产量"), cell(), cell("我司评估年产量"), cell(), cell("评估理由"), cell(colSpan = 3)
        )
        table.addAllCells(column)

        column = arrayOf(cell("项目预测说明(年销量、量产时间等)", 3), cell(colSpan = 7))
        table.addAllCells(column)

        column = arrayOf(
            cell("预计专项费用说明(不含税/万元)", 2, 2),
            cell("工装投入"),
            cell(colSpan = 2),
            cell("预计通用费用说明(不含税/万元)", 2, 2),
            cell("使用费用"),
            cell(colSpan = 2)
        )
        table.addAllCells(column)

        column = arrayOf(
            cell("开发费用"), cell(colSpan = 2), cell("其它"), cell(colSpan = 2)
        )
        table.addAllCells(column)

        column = arrayOf(cell("预计总费用(不含税/万元)", 3), cell(colSpan = 7))
        table.addAllCells(column)

        column = arrayOf(
            cell("产品报价(不含税/万元)", 3), cell(colSpan = 2), cell("报价单位(元/箱或元/托)", 2), cell(colSpan = 4)
        )
        table.addAllCells(column)

        column = arrayOf(
            cell("产品成本(不含税/万元)", 3), cell(colSpan = 2), cell("成本单位(元/箱或元/托)", 2), cell(colSpan = 4)
        )
        table.addAllCells(column)

        column = arrayOf(
            cell("客户承担方式", 2, 2), cell("非分摊方式承担(不含税/万元)", 2), cell(), cell("具体承担方式说明", 2), cell(colSpan = 4)
        )
        table.addAllCells(column)

        column = arrayOf(
            cell("分摊方式承担(不含税/万元)", 2), cell(), cell("数量(万箱/万托)", 2), cell(), cell("单位分摊价格"), cell()
        )
        table.addAllCells(column)

        column = arrayOf(
            cell("利润率(%)", 2), cell(colSpan = 3), cell("回收期(单位：年)", 3), cell(colSpan = 2)
        )
        table.addAllCells(column)

        column = arrayOf(
            cell("备注", 2, 2), cell(colSpan = 8, rowSpan = 2, height = 40f)
        )
        table.addAllCells(column)

        table.addColumnTable(
            approveColumn("制单人"),
            approveColumn("部门经理"),
            approveColumn("财务经理"),
            approveColumn("分管副总"),
            approveColumn("财务总监"),
            approveColumn("总经理")
        )
        document.add(table)
    }

    //扩展函数;为表格添加每一行单元格
    private fun PdfPTable.addColumnTable(vararg columns: Array<PdfPCell>) {
        columns.forEach { this.addAllCells(it) }
    }

    //审批行单元格集合;
    private fun approveColumn(label: String? = null): Array<PdfPCell> {
        return arrayOf(
            cell(label ?: "", 2, 2), cell(colSpan = 6, rowSpan = 2), cell(colSpan = 2, rowSpan = 2, height = 40f)
        )
    }

    //添加表格备注
    private fun addNotes(document: Document) {
        val table = PdfPTable(10)
        table.spacingBefore = 10f
        table.widthPercentage = 100f

        val note1 = "1、若多人一起出差，差旅费报销可由团队中职位最高的一个进行统一报销；若最高级别中有多人统计" + "（如各部门经理一起出差），则由统计中任意一名进行报销即可。此报销同时适用跨部门报销。"
        val note2 = "2、有超出标准的费用，请在申请报销时进行备注说明，如果没有进行说明，即使报销单被相关领导签批同意，" + "财务部仍有权拒绝报销。"
        val note3 =
            "3、差旅费用报销中，所有500元以上的消费，均需提供刷卡记录，如遇到无法刷卡服务的情况，请备注进行说明。" + "此外，500元以上餐旅费报销需要提供菜单；三星级以上宾馆住宿需提供住宿清单；市内交通公交、轻轨、地铁优先考虑；"
        table.addAllCells(
            arrayOf(
                noteCell("备注：", rowSpan = 3).apply { verticalAlignment = Element.ALIGN_TOP },
                noteCell(note1, colSpan = 9),
                noteCell(note2, colSpan = 10),
                noteCell(note3, colSpan = 10)
            )
        )
        document.add(table)
    }

    //扩展函数;添加单元格到表格
    private fun PdfPTable.addAllCells(cells: Array<PdfPCell>) {
        cells.forEach { this.addCell(it) }
    }

}