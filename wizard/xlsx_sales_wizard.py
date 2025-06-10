from email.policy import default

from dateutil.utils import today

from odoo import models, fields, api
from odoo.exceptions import UserError, ValidationError


class XlsxSalesWizard(models.TransientModel):
    _name = 'sales.excel.wizard'
    _description = 'Sales Report Wizard'

    order_ids = fields.Many2many('sale.order', string='Sale Order')
    user_ids = fields.Many2many('res.users', string='Sales Person')
    product_ids = fields.Many2many('product.product', string='Product')
    from_date = fields.Date(string="Date", help="Specify date "
                                                      "of the report period.")
    to_date = fields.Date(string="End Date", help="Specify the end date of the "
                                                  "report period.")
    detailed = fields.Boolean(string="Detailed Report", default=True)
    partner_ids = fields.Many2many('res.partner', string='Customer')
    categ_ids = fields.Many2many('product.category', string='Product Category')
    state = fields.Selection([('draft', "Quotation"),
    ('sent', "Quotation Sent"),
    ('sale', "Sales Order"),
    ('cancel', "Cancelled")], readonly=False, store=True)

    # @api.onchange('from_date', 'to_date')
    def _get_data(self):
        # domain = [('state', '!=', 'cancel')]
        domain = []
        if self.state:
            domain.append(('state', '=', self.state))
        if self.from_date:
            domain.append(('date_order', '>=', self.from_date))
        if self.to_date:
            domain.append(('date_order', '<=', self.to_date))
        if self.order_ids:
            domain.append(('id', '=', self.order_ids.ids))
        if self.user_ids:
            domain.append(('user_id', '=', self.user_ids.ids))
        if self.product_ids:
            sale_order_lines = self.env['sale.order.line'].search([('product_id', 'in', self.product_ids.ids)])
            domain.append(('order_line', '=', sale_order_lines.ids))
        if self.partner_ids:
            domain.append(('partner_id', '=', self.partner_ids.ids))
        if self.categ_ids:
            self.product_ids = self.env['product.product'].search([('categ_id', 'in', self.categ_ids.ids)])
            sale_order_lines = self.env['sale.order.line'].search([('product_id', 'in', self.product_ids.ids)])
            domain.append(('order_line', '=', sale_order_lines.ids))

        if self.from_date or self.to_date or self.order_ids or self.user_ids or self.product_ids or self.partner_ids or self.categ_ids or self.state:
            self.order_ids = self.env['sale.order'].search(domain)
        else:
            self.order_ids = self.env['sale.order'].search([])

    @api.onchange('state')
    def empty_order_ids(self):
        self.order_ids = False

    def set_to_date(self):
        if self.from_date and not self.to_date:
            self.to_date = today()

    def XlsxSalesReport(self):
        self._get_data()
        self.set_to_date()
        detailed = str(self.detailed)
        if not self.order_ids:
            raise ValidationError("No data available for printing.")
        return {
            'type': 'ir.actions.act_url',
            'url': f'/sales/excel/report/{self.order_ids.ids}/{self.user_ids.ids}/{self.product_ids.ids}/{self.from_date}/{self.to_date}/{detailed}/{self.partner_ids}',
            'target': 'new'
        }




