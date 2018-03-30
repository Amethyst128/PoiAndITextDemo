package data

import java.time.LocalDate

/**
 * 项目计划导出数据类
 * - 将项目数据转换为导出数据
 *
 * @author
 * @date 2018-03-21
 */
data class PProject(
    var name: String? = "",
    var id: String? = "",
    var responsible: String? = "",
    var customer: String? = "",
    var phases: List<PPhase>? = null,
    var startTime: LocalDate? = null,
    var endTime: LocalDate? = null
)

data class PPhase(
    var name: String? = "", var id: String? = "", var tasks: List<PTask>? = null
)

data class PTask(
    var name: String? = "",
    var output: String? = "",
    var responsible: String? = "",
    var startTime: LocalDate? = null,
    var endTime: LocalDate? = null
)

fun initReportData(): PProject {
    val tasks1 = listOf(
        PTask("定位主攻的防爆膜结构", "《新产品开发申请单》", "M.Snow", LocalDate.of(2017, 8, 21), LocalDate.of(2017, 8, 24)),
        PTask("他社品的数据收集和分析", "《他社品分析评价表》", "M.Snow", LocalDate.of(2017, 8, 25), LocalDate.of(2017, 8, 26)),
        PTask("防爆膜用PET膜的性能", "《产品原材料性能表》", "M.Snow", LocalDate.of(2017, 8, 27), LocalDate.of(2017, 8, 31)),
        PTask("防爆膜用离型膜的性能", "《产品原材料性能表》", "M.Snow", LocalDate.of(2017, 9, 3), LocalDate.of(2017, 9, 12)),
        PTask("防爆膜用胶黏剂的性能", "《产品原材料性能表》", "M.Snow", LocalDate.of(2017, 9, 14), LocalDate.of(2017, 9, 17))
    )
    val tasks2 = listOf(
        PTask("PET光学级透明膜的调查与评估", "《原材料评价报告》", "M.Snow", LocalDate.of(2017, 9, 4), LocalDate.of(2017, 9, 12)),
        PTask("PET离型膜的调查与评估", "《原材料评价报告》", "M.Snow", LocalDate.of(2017, 9, 22), LocalDate.of(2017, 9, 30)),
        PTask("OCA胶黏剂的调查与评估", "《胶黏剂评价报告》", "M.Snow", LocalDate.of(2017, 10, 2), LocalDate.of(2017, 10, 12)),
        PTask("产品设计", "《实验计划书》", "M.Snow", LocalDate.of(2017, 10, 14), LocalDate.of(2017, 10, 18)),
        PTask("配方开发验证实验", "《开发实验报告》", "M.Snow", LocalDate.of(2017, 10, 20), LocalDate.of(2017, 10, 26))
    )
    val tasks3 = listOf(
        PTask("结构、配比的探讨和最终确认", "《中试单》", "M.Snow", LocalDate.of(2017, 8, 21), LocalDate.of(2017, 9, 12)),
        PTask("中试样品完成", "《中试单》、《试制总结》", "M.Snow", LocalDate.of(2017, 10, 11), LocalDate.of(2017, 10, 12)),
        PTask("中试样品检测", "《中试实验报告》", "M.Snow", LocalDate.of(2017, 10, 15), LocalDate.of(2017, 10, 19)),
        PTask("中试样品总结", "《试制总结》", "M.Snow", LocalDate.of(2017, 10, 21), LocalDate.of(2017, 10, 30)),
        PTask("相关资料完成", "报告、TDS、MSDS等", "M.Snow", LocalDate.of(2017, 11, 9), LocalDate.of(2017, 11, 12))
    )
    val tasks4 = listOf(
        PTask("中试品", "", "M.Snow", LocalDate.of(2017, 11, 15), LocalDate.of(2017, 11, 18)),
        PTask("大试品", "", "M.Snow", LocalDate.of(2017, 11, 21), LocalDate.of(2017, 11, 27))
    )
    val tasks5 = listOf(
        PTask("第一次试制", "《试制总结》", "M.Snow", LocalDate.of(2017, 8, 28), LocalDate.of(2017, 12, 12)),
        PTask("第二次试制", "《试制总结》", "M.Snow", LocalDate.of(2017, 12, 16), LocalDate.of(2018, 4, 12))
    )
    val phases = listOf(
        PPhase("M1(项目立项)", "P001", tasks1),
        PPhase("M2(实验室设计开发)", "P002", tasks2),
        PPhase("M3(中试实验)", "P003", tasks3),
        PPhase("M4(客户评估)", "P004", tasks4),
        PPhase("M5(大试试制)", "P005", tasks5)
    )
    return PProject(
        "某测试项目开发", "OPT-YE-201803-56", "M.Snow", "xx客户", phases, LocalDate.of(2017, 8, 19), LocalDate.of(2018, 4, 25)
    )
}