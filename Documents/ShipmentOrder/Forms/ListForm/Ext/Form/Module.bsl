&AtClient
Procedure SyncWithAppSheet(Command)
    
    ShowQueryBox(
        New NotifyDescription("SyncWithAppSheetContinue", ThisObject),
        "Sync all documents with AppSheet?",
        QuestionDialogMode.YesNo
    );
    
EndProcedure

&AtClient
Procedure SyncWithAppSheetContinue(Result, Parameters) Export
    
    If Result = DialogReturnCode.Yes Then
        SyncWithAppSheetAtServer();
    EndIf;
    
EndProcedure

&AtServer
Procedure SyncWithAppSheetAtServer()
    
    UpdateCount = 0;
    SkipCount = 0;
    ErrorCount = 0;
    
    // Get AppSheet Configuration
    AppID = Constants.AppID.Get();
    AppSheetAPIKey = Constants.AccessKey.Get();
    TableName = Constants.SheetName.Get();
    
    If IsBlankString(AppID) Or IsBlankString(AppSheetAPIKey) Or IsBlankString(TableName) Then
        Message("AppSheet configuration is not complete. Please check Constants.");
        Return;
    EndIf;
    
    // Build API request
    ResourcePath = "/api/v2/apps/" + AppID + "/tables/" + TableName + "/Find";
    
    // Create HTTPS connection
    HTTPConnection = New HTTPConnection("api.appsheet.com", , , , , 60, New OpenSSLSecureConnection);
    
    // Prepare HTTP request
    HTTPRequest = New HTTPRequest(ResourcePath);
    HTTPRequest.Headers.Insert("ApplicationAccessKey", AppSheetAPIKey);
    HTTPRequest.Headers.Insert("Content-Type", "application/json");
    HTTPRequest.SetBodyFromString("{}", TextEncoding.UTF8);
    
    Try
        // Send request to AppSheet
        HTTPResponse = HTTPConnection.Post(HTTPRequest);
        
        If HTTPResponse.StatusCode = 200 Then
            
            // Get response as text
            ResponseText = HTTPResponse.GetBodyAsString(TextEncoding.UTF8);
            
            // Process each record in the response
            ProcessAppSheetResponse(ResponseText, UpdateCount, SkipCount, ErrorCount);
            
            // Show results
            Message("Sync completed: " + UpdateCount + " updated, " + SkipCount + " skipped, " + ErrorCount + " errors");
            
            // Refresh the list
            Items.List.Refresh();
            
        Else
            ErrorText = HTTPResponse.GetBodyAsString(TextEncoding.UTF8);
            Message("AppSheet API Error " + HTTPResponse.StatusCode + ": " + ErrorText);
        EndIf;
        
    Except
        ErrorInfo = ErrorInfo();
        Message("Connection error: " + ErrorInfo.Description);
    EndTry;
    
EndProcedure

&AtServer
Procedure ProcessAppSheetResponse(ResponseText, UpdateCount, SkipCount, ErrorCount)
    
    // Remove outer brackets if exists
    CleanResponse = ResponseText;
    If Left(CleanResponse, 1) = "[" Then
        CleanResponse = Mid(CleanResponse, 2, StrLen(CleanResponse) - 2);
    EndIf;
    
    // Split by "},{" to get individual records
    Position = 1;
    RecordCount = 0;
    
    While Position < StrLen(CleanResponse) Do
        
        RecordCount = RecordCount + 1;
        
        // Find start of this record
        RecordStart = Position;
        If Mid(CleanResponse, RecordStart, 1) = "{" Then
            RecordStart = RecordStart + 1;
        EndIf;
        
        // Find end of this record - look for "},{"
        RecordEnd = StrFind(CleanResponse, "},{", SearchDirection.FromBegin, RecordStart);
        
        If RecordEnd = 0 Then
            // This is the last record - find the closing }
            RecordEnd = StrFind(CleanResponse, "}", SearchDirection.FromBegin, RecordStart);
            If RecordEnd = 0 Then
                RecordEnd = StrLen(CleanResponse);
            EndIf;
            
            // Extract last record
            RecordText = "{" + Mid(CleanResponse, RecordStart, RecordEnd - RecordStart + 1);
            ProcessSingleRecord(RecordText, UpdateCount, SkipCount, ErrorCount);
            Break;
        Else
            // Extract this record
            RecordText = "{" + Mid(CleanResponse, RecordStart, RecordEnd - RecordStart) + "}";
            ProcessSingleRecord(RecordText, UpdateCount, SkipCount, ErrorCount);
            
            // Move to next record
            Position = RecordEnd + 2; // Skip "},{"
        EndIf;
        
    EndDo;
    
    Message("Processed " + RecordCount + " records from AppSheet");
    
EndProcedure

&AtServer
Procedure ProcessSingleRecord(RecordText, UpdateCount, SkipCount, ErrorCount)
    
    // Extract document number
    DocNumPos = StrFind(RecordText, """Doküman Numarası"":""");
    If DocNumPos = 0 Then
        Return;
    EndIf;
    
    DocNumStart = DocNumPos + StrLen("""Doküman Numarası"":""");
    DocNumEnd = StrFind(RecordText, """", SearchDirection.FromBegin, DocNumStart);
    DocumentNumber = Mid(RecordText, DocNumStart, DocNumEnd - DocNumStart);
    
    // Extract status
    StatusPos = StrFind(RecordText, """Doküman Statüsü"":""");
    If StatusPos = 0 Then
        Return;
    EndIf;
    
    StatusStart = StatusPos + StrLen("""Doküman Statüsü"":""");
    StatusEnd = StrFind(RecordText, """", SearchDirection.FromBegin, StatusStart);
    NewStatus = Mid(RecordText, StatusStart, StatusEnd - StatusStart);
    
    // Find document in 1C
    Query = New Query;
    Query.Text = 
    "SELECT
    |    ShipmentOrder.Ref AS Ref,
    |    ShipmentOrder.Status AS CurrentStatus
    |FROM
    |    Document.ShipmentOrder AS ShipmentOrder
    |WHERE
    |    ShipmentOrder.Number = &DocumentNumber";
    
    Query.SetParameter("DocumentNumber", DocumentNumber);
    QueryResult = Query.Execute();
    
    If QueryResult.IsEmpty() Then
        ErrorCount = ErrorCount + 1;
        Return;
    EndIf;
    
    Selection = QueryResult.Select();
    Selection.Next();
    
    // Check if status changed
    CurrentStatusString = String(Selection.CurrentStatus);
    
    If CurrentStatusString = NewStatus Then
        SkipCount = SkipCount + 1;
        Return;
    EndIf;
    
    // Update document
    Try
        DocumentObject = Selection.Ref.GetObject();
        
        // Set new status
        StatusChanged = False;
        
        If NewStatus = "NEW" Then
            DocumentObject.Status = Enums.ShipmentOrderStatus.NEW;
            StatusChanged = True;
        ElsIf NewStatus = "SHIPPED" Then
            DocumentObject.Status = Enums.ShipmentOrderStatus.SHIPPED;
            StatusChanged = True;
        ElsIf NewStatus = "DELIVERED" Then
            DocumentObject.Status = Enums.ShipmentOrderStatus.DELIVERED;
            StatusChanged = True;
        ElsIf NewStatus = "CANCELLED" Then
            DocumentObject.Status = Enums.ShipmentOrderStatus.CANCELLED;
            StatusChanged = True;
        ElsIf NewStatus = "PROCESSING" Then
            DocumentObject.Status = Enums.ShipmentOrderStatus.PROCESSING;
            StatusChanged = True;
        Else
            SkipCount = SkipCount + 1;
            Return;
        EndIf;
        
        If StatusChanged Then
            DocumentObject.Write();
            UpdateCount = UpdateCount + 1;
        EndIf;
        
    Except
        ErrorCount = ErrorCount + 1;
    EndTry;
    
EndProcedure

