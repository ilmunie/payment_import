# -*- encoding: utf-8 -*-
##############################################################################
#
#    OpenERP, Open Source Management Solution
#    Copyright (c) 2017 Eynes (http://www.eynes.com.ar)
#    All Rights Reserved. See AUTHORS for details.
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU Affero General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU Affero General Public License for more details.
#
#    You should have received a copy of the GNU Affero General Public License
#    along with this program.  If not, see <http://www.gnu.org/licenses/>.
#
##############################################################################

import base64
import os

import xlrd
from odoo import _, exceptions, fields, models, api

class AccontMover(models.Model):
    _inherit = "account.move"
    aux_payment_date = fields.Date()
    aux_journal_id = fields.Many2one('account.journal')
class paymentinfoImport(models.TransientModel):
    _name = "paymentinfo.import"
    _description = "paymentinfo Import"

    file = fields.Binary(string='File', filename="filename")
    filename = fields.Char(string='Filename', size=256)
    state = fields.Selection([('draft','Draft'),('done','Done')], string='State', default='draft')
    message = fields.Html(string='Message', readonly=True)

    def save_file(self, name, value):

        #path = os.path.abspath(os.path.dirname(__file__))
        path = '/tmp/%s' % name
        f = open(path, 'wb+')
        try:
            f.write(base64.decodebytes(value))
        finally:
            f.close()

        return path

    def _read_cell(self, sheet, row, cell):
        cell_type = sheet.cell_type(row, cell)
        cell_value = sheet.cell_value(row, cell)

        if cell_type in [1, 2]:  # 2: select, 1: text, 0: empty, 3: date
            return cell_value
        elif cell_type == 3:  # 3: date
            # Convertir el número de fecha en una fecha legible
            return xlrd.xldate.xldate_as_datetime(cell_value, 0)
        elif cell_type == 0:
            return None

        raise exceptions.UserError(_('Formato de archivo inválido'))

    def _set_default_warn_msg(self):
        msg = _("La importación se ha completado correctamente.")
        return """
        <td height="45" align="left">
            <strong>
                <span style='font-size:10.0pt;font-family:Arial;color:green'>%s</span>
            </strong>
        </td>
        """ % msg

    def _set_default_error_msg(self, message):
        msg = _("Se completó la importación con errores.")
        return """
        <td height="45" align="left">
            <strong>
                <span style='font-size:10.0pt;font-family:Arial;color:red'>""" + msg + """</span>
            </strong>
            <br>
            <strong>
                <span style='font-size:10.0pt;font-family:Arial;color:black'>""" + message[:-2] + """</span>
            </strong>
        </td>
        """


    def read_file(self):
        product = self.env['product.product']
        message = []
        rows_with_payments = []
        import pdb; pdb.set_trace()
        #path = os.path.abspath(os.path.dirname(__file__))
        path = '/tmp/%s' % self.filename
        # ~ f = open(path, 'a')

        path = self.save_file(self.filename, self.file)
        import pdb

        if path:
            book = xlrd.open_workbook(path)
            invoices = []
            sheet = book.sheets()[0]
            for curr_row in range(1, sheet.nrows):
                vals = {}
                invoice_name = str(self._read_cell(sheet, curr_row, 0) or '')
                if not invoice_name:
                    message.append(f'<span>FILA {curr_row}: ERROR FALTA NUMERO FACTURA</span>')
                    continue
                else:
                    searched_inv = self.env['account.move'].search([('ref','=',invoice_name),('move_type','=','in_invoice')])
                    if searched_inv:
                        invoice_id = searched_inv[0].id
                    else:
                        continue
                date_payment = self._read_cell(sheet, curr_row, 1) or False
                if date_payment:
                    payment_method = self._read_cell(sheet, curr_row, 2) or False
                    journal_id = False
                    #pdb.set_trace()
                    if not payment_method:
                        journal_id = 7
                    else:
                        cc = payment_method
                        if cc:
                            rp_bank = self.env['res.partner.bank'].search([('acc_number','=',cc)])
                            if rp_bank:
                                journal_ids = self.env['account.journal'].search([('bank_account_id','=',rp_bank[0].id)])
                                if journal_ids:
                                    journal_id = journal_ids[0].id
                    if journal_id:
                        vals['invoice_id'] = invoice_id
                        vals['aux_payment_date'] = date_payment
                        vals['aux_journal_id'] = journal_id
                    else:
                        message.append(f'<span>FILA {str(curr_row)} : ERROR DIARIO PAGO</span>')
                        continue
                invoices.append(vals)
            error = False
            for mess in message:
                if 'ERROR' in mess:
                    error = True
            if not error:
                for inv_dict in invoices:
                    inv = self.env['account.move'].search(inv_dict['invoice_id'])
                    payment = self.env['account.payment'].create({
                        'date': inv_dict['aux_payment_date'],
                        'payment_type': 'outbound',
                        'partner_type': 'supplier',
                        'partner_id': inv.partner_id.id,
                        'amount': abs(inv.amount_total_signed),
                        'currency_id': self.env.user.company_id.currency_id.id,
                        'journal_id': inv_dict['aux_journal_id'],
                    })
                    payment.action_post()
                    receivable_line = payment.line_ids.filtered('debit')
                    inv.js_assign_outstanding_line(receivable_line.id)
            self.state = 'done'
            view = self.env.ref('roc_custom.view_paymentinfo_import')
            return {
                'name': _('paymentinfo Import'),
                'res_model': 'paymentinfo.import',
                'type': 'ir.actions.act_window',
                'views': [(view.id, 'form')],
                'view_id': view.id,
                'target': 'new',
                'res_id': self.id,}


