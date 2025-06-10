import io
from ast import literal_eval
import xlsxwriter
from odoo import http
from odoo.http import request


class XlsxSalesReport(http.Controller):
    @http.route('/sales/excel/report/<string:order_ids>/<string:user_ids>/<string:product_ids>/<string:from_date>/<string:to_date>/<string:detailed>/<string:partner_id>', type='http', auth='user')
    def download_sales_excel_report(self, order_ids, from_date, to_date, detailed):
        order_ids = literal_eval(order_ids)
        # from_date = literal_eval(from_date)
        # to_date = literal_eval(to_date)
        detailed = True if detailed == 'True' else False

        output = io.BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Sales order')

        header_format = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1, 'align': 'center'})
        string_format = workbook.add_format({'border': 1, 'align': 'center'})
        price_format = workbook.add_format({'num_format': '$##,###00.00', 'border': 1, 'align': 'center'})
        t_string_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
        t_price_format = workbook.add_format({'bold': True, 'num_format': '$##,###00.00', 'border': 1, 'align': 'center', 'bg_color': '#D3D3D3'})
        head = workbook.add_format({'align': 'center', 'bold': True, 'font_size': '20px'})

        if not detailed:
            headers = ['Order', 'Creation Date', 'Delivery Date', 'Expected Date', 'Customer', 'Salesperson', 'Activities', 'Sales Team', 'Company', 'Untaxed Amount', 'Taxes', 'Total', 'Tags', 'Status', 'Invoice Status', 'To Invoice', 'Customer Reference', 'Expiration', 'Products']
            col_width = [20, 35, 20, 30, 25, 30, 30, 15, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20]
        else:
            headers = ['Order', 'Date','Customer', 'Product', 'Qty', 'Cost', 'Unit Price', 'Tax%', 'Taxes', 'Disc.%', 'Discount', 'Total', 'Tax incl.', 'Analytic Distribution', 'List Price', 'Profit', 'Margin', 'Price']
            col_width = [20, 20, 20, 40, 15, 20, 20, 15, 20, 15, 20, 20, 20, 20, 20, 20, 20, 20]

        for col_num, header in enumerate(headers):
            worksheet.write(4, col_num, header, header_format)
            worksheet.set_column(col_num, col_num, col_width[col_num])

            # TEst
        if order_ids:
            sale_ids = request.env['sale.order'].search([('id', 'in', order_ids)])
        else:
            sale_ids = request.env['sale.order'].search([])

        row_num = 5
        worksheet.merge_range('G1:N2', 'Sales Profit Report', head)
        if from_date and to_date :
            worksheet.merge_range('H3:M3', f"From {from_date} to {to_date}", head)

        t_qty = 0
        t_cost = 0
        t_taxes = 0
        t_disc = 0
        st_price = 0
        t_price = 0
        t_price1 = 0
        t_profit = 0
        t_margin = 0

        for sale in sale_ids:
            if not detailed:
                tag_names = ", ".join(tag.name for tag in sale.tag_ids)
                product_names = ", ".join(line.product_id.name for line in sale.order_line)
                worksheet.write(row_num, 0, sale.name if sale.name else '', string_format)
                worksheet.write(row_num, 1, sale.date_order.strftime('%Y-%m-%d %H:%M') if sale.date_order else '', string_format)
                worksheet.write(row_num, 2, sale.commitment_date.strftime('%Y-%m-%d') if sale.commitment_date else '', string_format)
                worksheet.write(row_num, 3, sale.expected_date.strftime('%Y-%m-%d %H:%M') if sale.expected_date else '', string_format)
                worksheet.write(row_num, 4, sale.partner_id.name if sale.partner_id else '', string_format)
                worksheet.write(row_num, 5, sale.user_id.name if sale.user_id else '', string_format)
                worksheet.write(row_num, 6, sale.activity_ids.display_name if sale.activity_ids else '', string_format)
                worksheet.write(row_num, 7, sale.team_id.name if sale.team_id else '', string_format)
                worksheet.write(row_num, 8, sale.company_id.name if sale.company_id else '', string_format)
                worksheet.write(row_num, 9, sale.amount_untaxed if sale.amount_untaxed else '', price_format)
                worksheet.write(row_num, 10, sale.amount_tax if sale.amount_tax else '', price_format)
                worksheet.write(row_num, 11, sale.amount_total if sale.amount_total else '', price_format)
                worksheet.write(row_num, 12, tag_names if tag_names else '', string_format)
                worksheet.write(row_num, 13, sale.state if sale.state else '', string_format)
                worksheet.write(row_num, 14, sale.invoice_status if sale.invoice_status else '', string_format)
                worksheet.write(row_num, 15, sale.amount_to_invoice if sale.amount_to_invoice else '', price_format)
                worksheet.write(row_num, 16, sale.client_order_ref if sale.client_order_ref else '', string_format)
                worksheet.write(row_num, 17, sale.validity_date.strftime('%Y-%m-%d') if sale.validity_date else '', string_format)
                worksheet.write(row_num, 18, product_names if product_names else '', string_format)
                row_num += 1
            else:
                for lines in sale.order_line:
                    profit = round(lines.product_id.list_price - lines.product_id.standard_price,2)
                    if lines.product_id.standard_price != 0:
                        margin = round((profit * 100) / lines.product_id.standard_price, 2)
                    worksheet.write(row_num, 0, sale.name if sale.name else '', string_format)
                    worksheet.write(row_num, 1, sale.date_order.strftime('%Y-%m-%d %H:%M') if sale.date_order else '', string_format)
                    worksheet.write(row_num, 2, sale.partner_id.name if sale.partner_id else '', string_format)
                    worksheet.write(row_num, 3, lines.name if sale.name else '', string_format)
                    worksheet.write(row_num, 4, lines.product_uom_qty if lines.product_uom_qty else '', string_format)
                    worksheet.write(row_num, 5, lines.product_id.standard_price if lines.product_id.standard_price else 0, price_format)
                    worksheet.write(row_num, 6, lines.price_unit if lines.price_unit else '', price_format)
                    worksheet.write(row_num, 7, lines.tax_id.name if lines.tax_id else '', string_format)
                    worksheet.write(row_num, 8, (lines.tax_id.amount/100)*lines.price_subtotal if lines.tax_id else '', price_format)
                    worksheet.write(row_num, 9, str(int(lines.discount))+'%' if lines.discount else '', string_format)
                    worksheet.write(row_num, 10, (lines.discount/100)*lines.price_unit if lines.discount else '', price_format)
                    worksheet.write(row_num, 11, lines.price_subtotal if lines.price_subtotal else '', price_format)
                    worksheet.write(row_num, 12, lines.price_total if lines.price_total else '', price_format)
                    worksheet.write(row_num, 13, lines.analytic_distribution if lines.analytic_distribution else '', price_format)
                    worksheet.write(row_num, 14, lines.product_id.list_price if lines.product_id.list_price else '', price_format)
                    worksheet.write(row_num, 15, profit if profit else '', price_format)
                    worksheet.write(row_num, 16, margin if profit != 0 else '', price_format)
                    worksheet.write(row_num, 17, lines.price_unit * lines.product_uom_qty if lines.price_unit else '', price_format)

                    t_qty += lines.product_uom_qty
                    # t_cost += lines.product_id.standard_price
                    t_cost += lines.product_id.standard_price * lines.product_uom_qty
                    t_taxes += (lines.tax_id.amount / 100) * lines.price_subtotal
                    t_disc += (lines.discount / 100) * lines.price_unit
                    st_price += lines.price_subtotal
                    t_price += lines.price_total
                    t_price1 += lines.price_total
                    t_profit += profit * lines.product_uom_qty
                    t_margin += margin * lines.product_uom_qty
                    row_num += 1
            worksheet.write(row_num, 4, t_qty if t_qty>0 else '', t_string_format)
            worksheet.write(row_num, 5, t_cost if t_cost>0 else 0,t_price_format)
            worksheet.write(row_num, 8, t_taxes if t_taxes>0 else '',t_price_format)
            worksheet.write(row_num, 10, t_disc if t_disc>0 else '',t_price_format)
            worksheet.write(row_num, 11, st_price if st_price>0 else '', t_price_format)
            worksheet.write(row_num, 12, t_price if t_price>0 else '', t_price_format)
            worksheet.write(row_num, 14, t_price1 if t_price1>0 else '', t_price_format)
            worksheet.write(row_num, 15, t_profit if t_profit>0 else '', t_price_format)
            worksheet.write(row_num, 16, t_margin if t_margin>0 else '', t_price_format)

            # worksheet.write(1, 3, 'Yes' if property.garden else 'No')

        workbook.close()
        output.seek(0)
        file_name = 'sales Report.xlsx'
        return request.make_response(
            output.getvalue(),
            headers=[
                ('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
                ('Content-Disposition', f'attachment; filename={file_name}')
            ]
        )
