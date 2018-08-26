# InvoiceProcessor
Python developed system for processing firms excel accounts and dynamically outputting invoices for given month
This system is entrusted with the following:

* On initialization automatically registers the current month and produces accounts according to that date.

* Scraping data from existing invoices updating associated client files accordingly

* A measure to prevent exceptional circumstances in transactions from affecting regular order data.

* Using client files to generate invoices for given month. 

* A measure to register exceptional dates such as a bank holiday and dynamically respond. 

* Such responses includes measures to analyse weeks purchases and adjust location of orders so as to minimize volume of trade affected

* System is configured in order that users within firm with no programming knowledge may access and make changes to accounts (and thus client files) solely through excel.
