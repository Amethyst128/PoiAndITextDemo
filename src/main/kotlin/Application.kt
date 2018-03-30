import data.initReportData
import form.QuoteApprovalForm
import reports.BusinessPlan

fun main(args: Array<String>) {
    println("form export start.")
    QuoteApprovalForm().createForm()
    println("form export end.")

    println("project report export start.")
    BusinessPlan().reportToFile(initReportData())
    println("project report end.")
}