CRM_ExchangeRates
=================
DataBridge that updates MS SQL Database using OpenExchangeRates.org API

MS SQL Database Tables:

tblCurrencies
tblCurrencyExchangeRates

tblDataBridgeLog
tblExchangeRatesLog

Configuration:
==============
Need to set values in app.config file:

1) www.openexchangerates.: API Key, UserName and Password
2) SMTP Server ServerName, Port, Username, Password, EmailFrom, EmailTo
3) SQL Server Username, Password and Database

