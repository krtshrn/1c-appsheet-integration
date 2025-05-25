
// Send ShipmentOrder document to AppSheet
Function SendShipmentOrderToAppSheet(DocumentObject) Export
	
	// Get AppSheet Configuration from Constants
	AppID = Constants.AppID.Get();
	AppSheetAPIKey = Constants.AccessKey.Get();
	TableName = Constants.SheetName.Get();
	
	If IsBlankString(AppID) Or IsBlankString(AppSheetAPIKey) Or IsBlankString(TableName) Then
		ErrorMessage = "AppSheet configuration is not complete. Please check Constants: AppID, AccessKey, SheetName";
		WriteLogEvent("AppSheet.Integration", EventLogLevel.Error, ErrorMessage);
		Return ErrorMessage;
	EndIf;
	
	ResourcePath = "/api/v2/apps/" + AppID + "/tables/" + TableName + "/add";
	
	// Create JSON body
	JSONWriter = New JSONWriter();
	JSONWriter.SetString();
	JSONWriter.WriteStartObject();
	JSONWriter.WritePropertyName("Action");
	JSONWriter.WriteValue("Add");
	JSONWriter.WritePropertyName("Properties");
	JSONWriter.WriteStartObject();
	JSONWriter.WritePropertyName("Locale");
	JSONWriter.WriteValue("tr-TR");
	JSONWriter.WritePropertyName("Timezone");
	JSONWriter.WriteValue("Europe/Istanbul");
	JSONWriter.WriteEndObject();
	JSONWriter.WritePropertyName("Rows");
	JSONWriter.WriteStartArray();
	JSONWriter.WriteStartObject();
	
	// Map 1C fields to AppSheet columns
	JSONWriter.WritePropertyName("Doküman Numarası");
	JSONWriter.WriteValue(DocumentObject.Number);
	JSONWriter.WritePropertyName("Tarih");
	JSONWriter.WriteValue(Format(DocumentObject.Date, "DF=dd.MM.yyyy"));
	JSONWriter.WritePropertyName("Personel");
	JSONWriter.WriteValue(String(DocumentObject.Employee));
	JSONWriter.WritePropertyName("Araç Plakası");
	JSONWriter.WriteValue(DocumentObject.VehicleLicensePlate);
	JSONWriter.WritePropertyName("Ambar");
	JSONWriter.WriteValue(String(DocumentObject.Warehouse));
	JSONWriter.WritePropertyName("Doküman Statüsü");
	JSONWriter.WriteValue(String(DocumentObject.Status));
	JSONWriter.WritePropertyName("Müşteri");
	JSONWriter.WriteValue(String(DocumentObject.Customer));
	JSONWriter.WritePropertyName("Adres");
	JSONWriter.WriteValue(String(DocumentObject.Address));
	
	JSONWriter.WriteEndObject();
	JSONWriter.WriteEndArray();
	JSONWriter.WriteEndObject();
	JSONBodyString = JSONWriter.Close();
	
	// Create HTTPS connection
	HTTPConnection = New HTTPConnection("api.appsheet.com", , , , , 60, New OpenSSLSecureConnection);
	
	// Prepare HTTP request
	HTTPRequest = New HTTPRequest(ResourcePath);
	HTTPRequest.Headers.Insert("ApplicationAccessKey", AppSheetAPIKey);
	HTTPRequest.Headers.Insert("Content-Type", "application/json");
	HTTPRequest.Headers.Insert("Accept", "application/json");
	HTTPRequest.SetBodyFromString(JSONBodyString, TextEncoding.UTF8);
	
	Try
		// Send POST request
		HTTPResponse = HTTPConnection.Post(HTTPRequest);
		
		If HTTPResponse.StatusCode = 200 Then
			// Success
			WriteLogEvent("AppSheet.Integration", 
			EventLogLevel.Information,
			"ShipmentOrder " + DocumentObject.Number + " successfully sent to AppSheet");
			Return "Success";
		Else
			// Error
			ErrorMessage = "HTTP Error " + HTTPResponse.StatusCode + ": " + HTTPResponse.GetBodyAsString(TextEncoding.UTF8);
			WriteLogEvent("AppSheet.Integration", 
			EventLogLevel.Error,
			"Failed to send ShipmentOrder " + DocumentObject.Number + ": " + ErrorMessage);
			Return ErrorMessage;
		EndIf;
		
	Except
		// Connection error
		ErrorInfo = ErrorInfo();
		ErrorMessage = "Connection error: " + ErrorInfo.Description;
		WriteLogEvent("AppSheet.Integration", 
		EventLogLevel.Error,
		"Connection error for ShipmentOrder " + DocumentObject.Number + ": " + ErrorMessage);
		Return ErrorMessage;
	EndTry;
	
EndFunction

