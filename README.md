# 1C:Enterprise & Google AppSheet Entegrasyonu

Bu proje, 1C:Enterprise 8.5 ile Google AppSheet arasında çift yönlü veri senkronizasyonunu sağlar. ShipmentOrder (Sevkiyat Siparişi) dokümanları üzerinden entegrasyon örneği sunulmuştur.

## Özellikler

- ✅ 1C'den AppSheet'e otomatik veri gönderimi
- ✅ AppSheet'ten 1C'ye manuel status senkronizasyonu
- ✅ Detaylı hata yönetimi ve loglama
- ✅ Toplu güncelleme desteği
- ✅ JSON tabanlı REST API kullanımı

## Gereksinimler

- 1C:Enterprise 8.5 veya üzeri
- Google AppSheet hesabı ve API erişimi
- SSL/HTTPS desteği

## Kurulum

### 1. AppSheet Tarafı

1. AppSheet'te yeni bir uygulama oluşturun
2. `ShipmentOrder` tablosu ekleyin ve şu alanları oluşturun:
   - Doküman Numarası (Text)
   - Tarih (Text)
   - Personel (Text)
   - Araç Plakası (Text)
   - Ambar (Text)
   - Doküman Statüsü (Text)

3. **Settings → Integrations → IN** bölümünden API'yi etkinleştirin
4. API Key'i kopyalayın

### 2. 1C:Enterprise Tarafı

#### Constants (Sabitler) Oluşturma

Configuration'da üç adet Constant oluşturun:

```
- AppID (String) - AppSheet uygulama ID'si
- AccessKey (String) - AppSheet API anahtarı  
- SheetName (String) - Tablo adı (örn: "ShipmentOrder")
```

#### Common Module Oluşturma

1. `AppSheetIntegration` adında bir Common Module oluşturun
2. **Server** ve **External connection** checkboxlarını işaretleyin
3. `SendShipmentOrderToAppSheet` fonksiyonunu ekleyin

#### Document Oluşturma

ShipmentOrder dokümanı için gerekli alanlar:
- Number (String)
- Date (Date)
- Employee (CatalogRef.Employees)
- VehicleLicensePlate (String)
- Warehouse (CatalogRef.Warehouses)
- Status (EnumRef.ShipmentOrderStatus)

Status Enum değerleri:
- NEW
- PROCESSING
- SHIPPED
- DELIVERED
- CANCELLED

## Kod Örnekleri

### Common Module - AppSheetIntegration

```1c
Function SendShipmentOrderToAppSheet(DocumentObject) Export
    
    // AppSheet konfigürasyonunu al
    AppID = Constants.AppID.Get();
    AppSheetAPIKey = Constants.AccessKey.Get();
    TableName = Constants.SheetName.Get();
    
    // API endpoint
    ResourcePath = "/api/v2/apps/" + AppID + "/tables/" + TableName + "/add";
    
    // JSON body oluştur
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
    
    // Alanları eşleştir
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
    
    JSONWriter.WriteEndObject();
    JSONWriter.WriteEndArray();
    JSONWriter.WriteEndObject();
    JSONBodyString = JSONWriter.Close();
    
    // HTTPS bağlantısı oluştur
    HTTPConnection = New HTTPConnection("api.appsheet.com", , , , , 60, New OpenSSLSecureConnection);
    
    // HTTP isteği hazırla
    HTTPRequest = New HTTPRequest(ResourcePath);
    HTTPRequest.Headers.Insert("ApplicationAccessKey", AppSheetAPIKey);
    HTTPRequest.Headers.Insert("Content-Type", "application/json");
    HTTPRequest.SetBodyFromString(JSONBodyString, TextEncoding.UTF8);
    
    Try
        // POST isteği gönder
        HTTPResponse = HTTPConnection.Post(HTTPRequest);
        
        If HTTPResponse.StatusCode = 200 Then
            Return "Success";
        Else
            ErrorMessage = "HTTP Error " + HTTPResponse.StatusCode + ": " + HTTPResponse.GetBodyAsString(TextEncoding.UTF8);
            Return ErrorMessage;
        EndIf;
        
    Except
        ErrorInfo = ErrorInfo();
        ErrorMessage = "Connection error: " + ErrorInfo.Description;
        Return ErrorMessage;
    EndTry;
    
EndFunction
```

### Document Module - Otomatik Gönderim

```1c
&AtServer
Procedure AfterWriteAtServer(CurrentObject, WriteParameters)

	Result = AppSheetIntegration.SendShipmentOrderToAppSheet(Object);
	
	If Result <> "Success" Then
		Message("Warning: Failed to send data to AppSheet: " + Result, MessageStatus.Important);
	Else
		Message("Shipment order #" + CurrentObject.Number + " was successfully sent to AppSheet.");
	EndIf;   
	
EndProcedure
```

## AppSheet'ten 1C'ye Senkronizasyon

DocumentList formuna eklenen Sync butonu ile AppSheet'teki değişiklikler 1C'ye çekilir.

### Nasıl Çalışır?

1. Sync butonuna tıklayın
2. Sistem AppSheet'ten tüm kayıtları çeker
3. Her kaydın status'ünü kontrol eder
4. Farklılık varsa 1C'deki kaydı günceller
5. Sonuç özeti gösterilir

## API İstek/Yanıt Örnekleri

### İstek (1C → AppSheet)
```json
{
  "Action": "Add",
  "Properties": {
    "Locale": "tr-TR",
    "Timezone": "Europe/Istanbul"
  },
  "Rows": [{
    "Doküman Numarası": "000000001",
    "Tarih": "24.05.2025",
    "Personel": "Hasan Kaya",
    "Araç Plakası": "34AA34",
    "Ambar": "GEBZE",
    "Doküman Statüsü": "NEW"
  }]
}
```

### Yanıt (AppSheet → 1C)
```json
[{
  "_RowNumber": "2",
  "Doküman Numarası": "000000001",
  "Tarih": "24.05.2025",
  "Personel": "Hasan Kaya",
  "Araç Plakası": "34AA34",
  "Ambar": "Gebze",
  "Doküman Statüsü": "DELIVERED"
}]
```

