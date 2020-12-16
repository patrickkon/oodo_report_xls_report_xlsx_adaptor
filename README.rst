.. image:: https://img.shields.io/badge/licence-AGPL--3-blue.svg
    :target: http://www.gnu.org/licenses/agpl-3.0-standalone.html
    :alt: License: AGPL-3

================
report_xls to report_xlsx Adaptor
================

Create .xls generated from odoo's report_xls library, to a .xlsx file, with minimal code changes.

==============
General Steps for Usage (in a given report_xls child class)
==============
Remove these 2 lines: ::

    import xlwt

    from openerp.addons.report_xls.report_xls import report_xls

Add these 2 lines: ::

    python from openerp.addons.report_xls_to_report_xlsx_adaptor.report.report_xlsx import ReportXlsx as report_xls

    xlwt = wb

==============
Installation and usage of report_xls and report_xlsx
==============

Refer to the individual packages: 

https://apps.odoo.com/apps/modules/10.0/report_xls/ 

https://github.com/OCA/reporting-engine/tree/12.0/report_xlsx


.. image:: https://odoo-community.org/website/image/ir.attachment/5784_f2813bd/datas
   :alt: Try me on Runbot
   :target: https://runbot.odoo-community.org/runbot/143/8.0

Bug Tracker
===========

Bugs are tracked on `GitHub Issues <https://github.com/patrickkon/oodo_report_xls_to_report_xlsx_adaptor/issues>`_.
In case of trouble, please check there if your issue has already been reported.
If you spotted it first, help us smashing it by providing a detailed and welcomed feedback
`here <https://github.com/OCA/reporting-engine/issues/new?body=module:%20report_xlsx%0Aversion:%208.0%0A%0A**Steps%20to%20reproduce**%0A-%20...%0A%0A**Current%20behavior**%0A%0A**Expected%20behavior**>`_.

Credits
=======

* Icon taken from http://www.icons101.com/icon/id_67712/setid_2096/Boxed_Metal_by_Martin/xlsx.

Contributors
------------

* Patrick Kon <patrikon2-c@my.cityu.edu.hk>

Maintainer
----------

.. image:: https://odoo-community.org/logo.png
   :alt: Odoo Community Association
   :target: https://odoo-community.org

This module is maintained by the OCA.

OCA, or the Odoo Community Association, is a nonprofit organization whose mission is to support the collaborative development of Odoo features and promote its widespread use.

To contribute to this module, please visit https://odoo-community.org.
