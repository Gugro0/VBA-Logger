# VBA-Logger

Logging for VBA Applications
----------------------------

Configurable and flexible Logging Framework for VBA Applications (Access,Word,Excel).
Collect your application errors, infos or traces with a simple API and save them to text files, database tables or send them by email.
Apply filter conditions on severity, source, number or any text.

You basic interface is :
  ` Logger.Log.[Error|Debug|Trace|Message] `

e.g logging an Error from the UI.Customer class in Procedure DisplayCustomer yo may use:

 ` Logger.Log.Error "UI", "Customer", "DisplayCustomer", 1234, "Unknown Customer ID" `
