# VisioSamples

## CallOneAddinFromAnother
An example illustrating calling one Visio VSTO addin (VisioAddin1) from another Visio addin (VisioAddin2)
The first add-in registers API by implementing `RequestComAddInAutomationService` and the second calls that API using `Application.COMAddIns("...").Object.`
