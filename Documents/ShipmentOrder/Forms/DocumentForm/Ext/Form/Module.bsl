
&AtServer
Procedure AfterWriteAtServer(CurrentObject, WriteParameters)

	Result = AppSheetIntegration.SendShipmentOrderToAppSheet(Object);
	
	If Result <> "Success" Then
		Message("Warning: Failed to send data to AppSheet: " + Result, MessageStatus.Important);
	Else
		Message("Shipment order #" + CurrentObject.Number + " was successfully sent to AppSheet.");
	EndIf;   
	
EndProcedure
