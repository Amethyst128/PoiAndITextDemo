package data

/**
 * 数据类;用来填充表单模板数据
 *
 * @author
 * @date 2018-02-03
 */
data class QuoteApproval(
    var title: String? = "",
    var managerCode: String? = "",
    var applyUser: String? = "",
    var owningGroup: String? = "",
    var owningRole: String? = "",
    var applyDate: String? = "",
    var customerName: String? = "",
    var productName: String? = "",
    var customerExpects: String? = "",
    var weExpects: String? = "",
    var evaluateReason: String? = "",
    var forecastDesc: String? = "",
    var toolInvestment: String? = "",
    var useCost: String? = "",
    var developmentCost: String? = "",
    var otherCost: String? = "",
    var estimatedTotalCost: String? = "",
    var productPricing: String? = "",
    var quotationUnit: String? = "",
    var productCost: String? = "",
    var costUnit: String? = "",
    var unaffordableWay: String? = "",
    var unaffordableExplain: String? = "",
    var assignWay: String? = "",
    var assignQuantity: String? = "",
    var unitSharePrice: String? = "",
    var profitRate: String? = "",
    var paybackPeriod: String? = "",
    var remarks: String? = "",
    var formOwner: String? = "",
    var formOwnerSignature: String? = "",
    var financialManager: String? = "",
    var financialManagerSignature: String? = "",
    var departmentManager: String? = "",
    var departmentManagerSignature: String? = "",
    var vicePresident: String? = "",
    var vicePresidentSignature: String? = "",
    var financialDirector: String? = "",
    var financialDirectorSignature: String? = "",
    var generalManager: String? = "",
    var generalManagerSignature: String? = ""
)