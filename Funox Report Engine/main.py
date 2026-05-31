import pandas as pd
from openpyxl import load_workbook
from docxtpl import DocxTemplate


client_name = "徐州某某有限公司"

excel_path = "data_cn.xlsx"
template_path = "template.docx"
output_path = "output.docx"


def get_visible_sheets(excel_path):
    """获取 Excel 中显示的 Sheet，不读取隐藏 Sheet"""
    wb = load_workbook(excel_path, read_only=True, data_only=True)

    visible_sheets = [
        ws.title
        for ws in wb.worksheets
        if ws.sheet_state == "visible"
    ]

    wb.close()
    return visible_sheets


def read_sheet(sheet_name):
    """读取 Sheet 内容"""
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    amount_columns = [
        "本期金额",
        "上期金额",
    ]

    for col in amount_columns:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda x: f"{float(x):,.2f}"
                if pd.notna(x) and str(x).strip() != ""
                else ""
            )

    df = df.fillna("")

    return df.to_dict("records")


# Excel Sheet 名称 与 Word 模板变量名的对应关系
sheet_map = {
    "货币资金": "cash_rows",
    "以公允价值计量且其变动计入当期损益的金融资产": "financial_assets_fvpl_rows",
    "应收票据": "notes_receivable_rows",
    "应收账款": "accounts_receivable_rows",
    "预付账款": "prepayments_rows",
    "其他应收款": "other_receivables_rows",
    "存货": "inventory_rows",
    "一年内到期的非流动资产": "non_current_assets_due_within_one_year_rows",
    "其他流动资产": "other_current_assets_rows",

    "长期股权投资": "long_term_equity_investments_rows",
    "投资性房地产": "investment_property_rows",
    "固定资产": "property_plant_equipment_rows",
    "在建工程": "construction_in_progress_rows",
    "无形资产": "intangible_assets_rows",
    "使用权资产": "right_of_use_assets_rows",
    "长期待摊费用": "long_term_deferred_expenses_rows",
    "递延所得税资产": "deferred_tax_assets_rows",
    "其他非流动资产": "other_non_current_assets_rows",

    "短期借款": "short_term_borrowings_rows",
    "以公允价值计量且其变动计入当期损益的金融负债": "financial_liabilities_fvpl_rows",
    "应付票据": "notes_payable_rows",
    "应付账款": "accounts_payable_rows",
    "预收账款": "advance_from_customers_rows",
    "应付职工薪酬": "employee_compensation_payable_rows",
    "应交税费": "taxes_and_surcharges_payable_rows",
    "其他应付款": "other_payables_rows",
    "一年内到期的非流动负债": "non_current_liabilities_due_within_one_year_rows",
    "其他流动负债": "other_current_liabilities_rows",

    "长期借款": "long_term_borrowings_rows",
    "应付债券": "bonds_payable_rows",
    "预计负债": "provisions_rows",
    "租赁负债": "lease_liabilities_rows",
    "递延收益": "deferred_income_rows",
    "递延所得税负债": "deferred_tax_liabilities_rows",
    "其他非流动负债": "other_non_current_liabilities_rows",

    "股本(实收资本)": "share_capital_rows",
    "资本公积": "capital_reserve_rows",
    "其他综合收益": "other_comprehensive_income_rows",

    "营业收入": "revenue_rows",
    "营业成本": "cost_of_sales_rows",
    "税金及附加": "taxes_and_surcharges_rows",
    "销售费用": "selling_expenses_rows",
    "管理费用": "administrative_expenses_rows",
    "财务费用": "financial_expenses_rows",
    "资产减值损失": "asset_impairment_loss_rows",
    "信用减值损失": "credit_impairment_loss_rows",
    "投资收益": "investment_income_rows",
    "其他收益": "other_gains_rows",
    "营业外收入": "non_operating_income_rows",
    "营业外支出": "non_operating_expenses_rows",
}


# 获取显示的 Sheet
visible_sheets = get_visible_sheets(excel_path)

print("当前显示的 Sheet：")
for sheet in visible_sheets:
    print("-", sheet)


# 初始化模板数据
context = {
    "client_name": client_name
}


# 根据 sheet_map 自动读取数据
for sheet_name, var_name in sheet_map.items():
    if sheet_name in visible_sheets:
        context[var_name] = read_sheet(sheet_name)
    else:
        context[var_name] = []


# 渲染 Word 模板
doc = DocxTemplate(template_path)
doc.render(context)
doc.save(output_path)

print(f"已生成 {output_path}")