package reports

import data.PProject
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.RegionUtil
import org.apache.poi.xssf.usermodel.XSSFCellStyle
import org.apache.poi.xssf.usermodel.XSSFFont
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.FileInputStream
import java.io.FileOutputStream
import java.io.IOException
import java.text.SimpleDateFormat
import java.time.LocalDate
import java.time.temporal.ChronoUnit
import java.util.*


/**
 * 项目报表
 *
 * @author
 * @date 2018-03-21
 */
class BusinessPlan {
    private val input = "temp/project.xlsx"
    private lateinit var workBook: XSSFWorkbook  //工作表
    private lateinit var sheet: Sheet  //页面

    /**
     * 导出到服务器本地
     *
     * @param project
     */
    fun reportToFile(project: PProject) {
        try {
            val out = FileOutputStream("src/files/report.xlsx")  //输出文件
            workBook = this.product(project)  //写数据
            workBook.write(out)
            out.flush()
            out.close()
        } catch (e: IOException) {
            e.printStackTrace()
            throw IOException("Export error.")
        } catch (e: Exception) {
            e.printStackTrace()
            throw Exception("Export error.")
        }
    }

    /**
     * 生成计划报表
     *
     * @param project
     * @return
     */
    private fun product(project: PProject): XSSFWorkbook {
        try {
            project.phases ?: throw Exception("Your Project not phases.")
            project.phases?.flatMap { it.tasks ?: listOf() } ?: throw Exception("Your project not tasks.")

            //获取模板
            val fis = FileInputStream(this::class.java.classLoader.getResource(input).path)
            //如果是xlsx，2007，用XSSF,如果是xls，2003，用HSSF
            workBook = XSSFWorkbook(fis)
            //获取第一个sheet页。
            sheet = workBook.getSheetAt(0)
            //设置空白区域
            sheet.isDisplayGridlines = false

            //拿到第三行,写入项目信息
            sheet.getRow(3).getCell(0).also {
                //创建样式
                val font = workBook.font().bold()
                it.cellStyle = workBook.border().font(font)
                //设值
                it.setCellValue("客户名称: ${project.customer}    课题名称: ${project.name}    项目编号: ${project.id}    责任人: ${project.responsible}")
            }

            val lastCell = this.setChart(project.startTime, project.endTime) //图头信息
            this.createContent(project, lastCell)  //写入正文及图表

            sheet.createFreezePane(7, 0, 7, 0)  //固定区域
        } catch (e: IOException) {
            e.printStackTrace()
            throw IOException("Export error.")
        } catch (e: Exception) {
            e.printStackTrace()
            throw Exception("Export error.")
        }
        return workBook
    }

    /**
     * 写入项目阶段及任务内容
     *
     * @param project
     * @param lastCell
     */
    private fun createContent(project: PProject, lastCell: Int) {
        val phases = project.phases
        val font = workBook.font() //自定义字体

        var lastRow = 6  //每个阶段完成后，所在的最后一列
        var line = 0  //任务计数，需要给任务编号
        phases?.forEachIndexed p@{ _, phase ->
            var temp = 0  //临时变量,用来计算合并的单元格
            val tasks = phase.tasks ?: return@p
            tasks.forEachIndexed t@{ _, task ->
                line = line.inc()
                val row = sheet.createRow(lastRow)
                row.createCell(1).also { it.cellStyle = workBook.border().font(font).alignment() }
                    .setCellValue(line.toDouble())
                row.createCell(2).also { it.cellStyle = workBook.border().font(font) }.setCellValue("${task.name}")
                row.createCell(3).also { it.cellStyle = workBook.border().font(font) }.setCellValue("${task.output}")
                row.createCell(4).also { it.cellStyle = workBook.border().font(font).alignment() }
                    .setCellValue("${task.responsible}")
                row.createCell(5).also { it.cellStyle = workBook.border().font(font).alignment() }
                    .setCellValue("${task.startTime}")
                row.createCell(6).also { it.cellStyle = workBook.border().font(font).alignment() }
                    .setCellValue("${task.endTime}")
                for (i in 7 until lastCell) {
                    row.createCell(i).also { it.cellStyle = workBook.border() }
                }
                val startDays =
                    ChronoUnit.DAYS.between(project.startTime, task.startTime).toInt() //获取任务开始日期和项目开始日期之间的天数差
                val endDays = ChronoUnit.DAYS.between(task.startTime, task.endTime).toInt()  //获取任务结束时间和任务开始时间之间的天数差
                if (startDays < 0 || endDays < 0) return@t

                val colorStyle = workBook.createCellStyle().randomColor()
                for (i in (startDays + 7)..(endDays + startDays + 7)) {
                    row.createCell(i).also { it.cellStyle = colorStyle }
                }

                lastRow += 1
                temp += 1
            }
            //拿到任务第一行,并设值
            val row = sheet.getRow(lastRow - tasks.size)
            row.createCell(0).also { it.cellStyle = workBook.border().font(font).center().wrapText() }
                .setCellValue("${phase.name}")
            //合并阶段区域
            sheet.addMergedRegion(CellRangeAddress(lastRow - temp, lastRow - 1, 0, 0).also {
                RegionUtil.setBorderBottom(BorderStyle.THIN, it, sheet)
            })
        }

        phases?.flatMap { it.tasks ?: listOf() }?.let {
            //设置底部信息
            val footer = sheet.createRow(it.size + 7)
            footer.createCell(2).also { it.cellStyle = workBook.createCellStyle().font(workBook.font()) }
                .setCellValue("编制:")
            footer.createCell(4).also { it.cellStyle = workBook.createCellStyle().font(workBook.font()) }
                .setCellValue("核准:")
        }
    }

    /**
     * 写入计划图表头部信息
     *
     * @param lastCell
     * @param startTime
     */
    private fun setChartHeader(lastCell: Int, startTime: Calendar) {
        val format = SimpleDateFormat("yyyy-MM-dd")
        //创建样式
        val style = workBook.createCellStyle().font(workBook.font())
        style.dataFormat = workBook.createDataFormat().getFormat("yyy-MM-dd HH:mm")
        style.setAlignment(HorizontalAlignment.LEFT)

        sheet.autoSizeColumn(0, false)
        //拿到第三行,写入时间
        val row3 = sheet.getRow(3)

        row3.createCell(7).also { it.cellStyle = style }.cellFormula = "NOW()"
        row3.createCell(26).also { it.cellStyle = style }.setCellValue("课题周期: ${format.format(startTime.time)} ")

        //合并当前时间区域{H~U}
        sheet.addMergedRegion(CellRangeAddress(3, 3, 7, 25).also {
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, it, sheet)
            RegionUtil.setBorderBottom(BorderStyle.THIN, it, sheet)
        })
        //合并课程周期区域{}
        sheet.addMergedRegion(CellRangeAddress(3, 3, 26, lastCell - 1).also {
            RegionUtil.setBorderTop(BorderStyle.MEDIUM, it, sheet)
            RegionUtil.setBorderRight(BorderStyle.THIN, it, sheet)
            RegionUtil.setBorderBottom(BorderStyle.THIN, it, sheet)
        })
    }

    /**
     * 写入计划图表信息
     *
     * @param startDate
     * @param endDate
     * @return
     */
    private fun setChart(startDate: LocalDate?, endDate: LocalDate?): Int {
        if (null == startDate || null == endDate || startDate == endDate) return 0

        val start = GregorianCalendar(startDate.year, startDate.monthValue - 1, startDate.dayOfMonth)
        val end = GregorianCalendar(endDate.year, endDate.monthValue - 1, endDate.dayOfMonth)

        //如果开始时间大于结束时间【说明时间错误,直接结束】
        if (end.before(start)) return 0

        return if (end[Calendar.YEAR] == start[Calendar.YEAR]) {
            sameYearWrite(start, end)
        } else {
            differentYearWrite(start, end)
        }
    }

    /**
     * 开始年和结束年不同时调用该方法【开始年不能大于结束年】
     *
     * @param start
     * @param end
     * @return
     */
    private fun differentYearWrite(start: GregorianCalendar, end: GregorianCalendar): Int {
        var lastCell = 7  //记录最后一个单元格
        var currentCell = lastCell  //定义当前单元格

        val startYear = start[Calendar.YEAR] //记录开始年
        val endYear = end[Calendar.YEAR]  //记录结束年
        //如果开始年小于结束年，计算年差和月差
        val yearSpace = endYear - startYear  //计算结束年和开始年间距
        var currentYear = startYear //设置当前年-1
        for (y in 0 until yearSpace + 1) {
            val temp = GregorianCalendar.getInstance()  //定义临时日历
            temp.set(Calendar.YEAR, currentYear)  //设置年为当前年

            if (startYear == currentYear) {  //如果是开始年，也就是当前年的话
                temp.set(Calendar.MONTH, start[Calendar.MONTH])  //设置开始月
                temp.set(Calendar.DAY_OF_MONTH, start[Calendar.DAY_OF_MONTH])  //设置开始天
                for (i in temp[Calendar.MONTH]..11) {  //从开始月到最大月
                    while (temp[Calendar.MONTH] == i) {
                        val day = temp[Calendar.DAY_OF_MONTH]

                        //写入单元格内容
                        sheet.getRow(5).createCell(currentCell).also {
                            it.cellStyle = workBook.border().font(workBook.font(8)).alignment()
                        }.setCellValue(day.toDouble())
                        sheet.setColumnWidth(currentCell, 2 * 256 + 20)

                        temp.add(Calendar.DAY_OF_MONTH, 1)
                        currentCell += 1  //下一个单元格
                    }
                    sheet.getRow(4).createCell(lastCell).also {
                        it.cellStyle = workBook.border().font(workBook.font()).center()
                    }.setCellValue("$currentYear 年 ${if (temp[Calendar.MONTH] == 0) 12 else temp[Calendar.MONTH]} 月")
                    sheet.addMergedRegion(CellRangeAddress(4, 4, lastCell, currentCell - 1).also {
                        RegionUtil.setBorderRight(BorderStyle.THIN, it, sheet)
                    })
                    lastCell = currentCell  //最后单元格
                }
                currentYear += 1  //年+1
                temp.clear()  //清空临时日历
            } else if (currentYear == endYear) {
                temp.set(Calendar.DAY_OF_MONTH, 1)  //设置开始天为1
                for (i in 0..end[Calendar.MONTH]) {
                    temp.set(Calendar.MONTH, i)  //设置月为i
                    while (temp[Calendar.MONTH] == i) {
                        val day = temp[Calendar.DAY_OF_MONTH]

                        //写入单元格内容
                        sheet.getRow(5).createCell(currentCell).also {
                            it.cellStyle = workBook.border().font(workBook.font(8)).alignment()
                        }.setCellValue(day.toDouble())
                        sheet.setColumnWidth(currentCell, 2 * 256 + 20)

                        temp.add(Calendar.DAY_OF_MONTH, 1)
                        currentCell += 1  //下一个单元格
                    }
                    sheet.getRow(4).createCell(lastCell).also {
                        it.cellStyle = workBook.border().font(workBook.font()).center()
                    }.setCellValue("$currentYear 年 ${if (temp[Calendar.MONTH] == 0) 12 else temp[Calendar.MONTH]} 月")
                    sheet.addMergedRegion(CellRangeAddress(4, 4, lastCell, currentCell - 1).also {
                        RegionUtil.setBorderRight(BorderStyle.THIN, it, sheet)
                    })
                    lastCell = currentCell  //最后单元格
                }
                currentYear += 1  //年+1
                temp.clear()  //清空临时日历
            } else if (currentYear != startYear && currentYear != endYear) {
                temp.set(Calendar.DAY_OF_MONTH, 1)  //设置开始天为1
                for (i in 0..11) {
                    temp.set(Calendar.MONTH, i)   //设置月i
                    while (temp[Calendar.MONTH] == i) {
                        val day = temp[Calendar.DAY_OF_MONTH]

                        //写入单元格内容
                        sheet.getRow(5).createCell(currentCell).also {
                            it.cellStyle = workBook.border().font(workBook.font(8)).alignment()
                        }.setCellValue(day.toDouble())
                        sheet.setColumnWidth(currentCell, 2 * 256 + 20)

                        temp.add(Calendar.DAY_OF_MONTH, 1)
                        currentCell += 1  //下一个单元格
                    }
                    sheet.getRow(4).createCell(lastCell).also {
                        it.cellStyle = workBook.border().font(workBook.font()).center()
                    }.setCellValue("$currentYear 年 ${if (temp[Calendar.MONTH] == 0) 12 else temp[Calendar.MONTH]} 月")
                    sheet.addMergedRegion(CellRangeAddress(4, 4, lastCell, currentCell - 1).also {
                        RegionUtil.setBorderRight(BorderStyle.THIN, it, sheet)
                    })
                    lastCell = currentCell  //最后单元格
                }
                currentYear += 1  //年+1
            }
        }
        this.setChartHeader(lastCell, start)  //设置甘特图头信息
        return lastCell
    }

    /**
     * 开始年和结束年相同时调用该方法
     *
     * @param start
     * @param end
     * @return
     */
    private fun sameYearWrite(start: GregorianCalendar, end: GregorianCalendar): Int {
        var lastCell = 7  //记录最后一个单元格
        //计算两个月份之差，如果结果不为0，则数量+1
        val monthSpace = end[Calendar.MONTH] - start[Calendar.MONTH]
        for (i in 0..if (start[Calendar.MONTH] != 0) monthSpace + 1 else monthSpace) {
            if (start[Calendar.MONTH] != i) continue  //如果不是当前月，直接跳过

            var currentCell = lastCell  //定义当前单元格
            while (start[Calendar.MONTH] == i) {
                val day = start[Calendar.DAY_OF_MONTH]
                //写入单元格内容
                sheet.getRow(5).createCell(currentCell).also {
                    it.cellStyle = workBook.border().font(workBook.font(8)).alignment()
                }.setCellValue(day.toDouble())  //toDouble的目的是【Excel默认是Double类型，而不是整形】
                sheet.setColumnWidth(currentCell, 2 * 256 + 20) //设置【天】单元格宽度

                start.add(Calendar.DAY_OF_MONTH, 1)
                currentCell += 1  //下一个单元格
            }
            sheet.getRow(4).createCell(lastCell).also {
                it.cellStyle = workBook.border().font(workBook.font()).center()
            }.setCellValue("${start[Calendar.YEAR]} 年 ${start[Calendar.MONTH]} 月")
            sheet.addMergedRegion(CellRangeAddress(4, 4, lastCell, currentCell - 1).also {
                RegionUtil.setBorderRight(BorderStyle.THIN, it, sheet)
            })

            lastCell = currentCell  //最后单元格
        }
        this.setChartHeader(lastCell, start)  //设置甘特图头信息
        return lastCell
    }

}

//-----------------------------------------------------------------扩展函数;可使用私有方法替代.
/**
 * 设置图表颜色【随机】
 */
private fun XSSFCellStyle.randomColor(): XSSFCellStyle {
    val colors = arrayOf(
        IndexedColors.DARK_RED,
        IndexedColors.GREY_40_PERCENT,
        IndexedColors.BLUE_GREY,
        IndexedColors.SKY_BLUE,
        IndexedColors.LIGHT_ORANGE,
        IndexedColors.INDIGO,
        IndexedColors.DARK_TEAL,
        IndexedColors.SEA_GREEN,
        IndexedColors.AQUA,
        IndexedColors.MAROON
    )
    this.fillForegroundColor = colors[(Math.random() * colors.size).toInt()].index
    this.setFillPattern(FillPatternType.SOLID_FOREGROUND)
    return this
}

/**
 * 设置默认边框(THIN)
 */
private fun XSSFWorkbook.border(): XSSFCellStyle {
    val style = this.createCellStyle()
    style.setBorderTop(BorderStyle.THIN)
    style.setBorderBottom(BorderStyle.THIN)
    style.setBorderLeft(BorderStyle.THIN)
    style.setBorderRight(BorderStyle.THIN)
    return style
}

/**
 * 设置自动换行
 */
private fun XSSFCellStyle.wrapText(): XSSFCellStyle {
    this.wrapText = true
    return this
}


/**
 * 设置单元格横向居中
 */
private fun XSSFCellStyle.alignment(): XSSFCellStyle {
    this.setAlignment(HorizontalAlignment.CENTER)
    return this
}

/**
 * 设置单元格居中
 */
private fun XSSFCellStyle.center(): XSSFCellStyle {
    this.setAlignment(HorizontalAlignment.CENTER)
    this.setVerticalAlignment(VerticalAlignment.CENTER)
    return this
}

/**
 * 设置单元格字体
 */
private fun XSSFCellStyle.font(font: Font): XSSFCellStyle {
    this.setFont(font)
    return this
}

/**
 * 创建Excel默认字体
 */
private fun XSSFWorkbook.font(size: Short = 10): XSSFFont {
    val font = this.createFont()
    font.fontName = "微软雅黑"
    font.fontHeightInPoints = size
    return font
}

/**
 * 字体加粗
 */
private fun XSSFFont.bold(): XSSFFont {
    this.bold = true
    return this
}
