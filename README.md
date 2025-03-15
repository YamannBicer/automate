Qlik Sense Hub linki üzerinden "Turkey_DSG_P&L" raporunun "MAIN TABLE REAL-TIME" sayfasını açıyoruz.
Seçili günün, bulunduğumuz günden bir önceki gün olduğundan emin oluyoruz.
Tablodaki verilerin totals kısmının 0'dan farklı olmasını bekliyoruz. (Financial Gain kolonu hariç)
Tablodaki verileri excel'e indiriyoruz.
Sharepoint'da tuttuğumuz TürkiyeDSGDailyReport.xlsx linkli template'e exceldeki verileri aktarıyoruz.
Results sheet'indeki tabloyu pdf formatında aşağıdaki metni de ekleyerek listeli kişilere iletiyoruz.

"TR-Intraday (PURE)" <tr-intraday@pureenergy.com.tr>
"OPERASYON" <operasyon@pureenergy.com.tr>
"Origination TR (PURE)" <originationtr@pureenergy.com.tr>

To: burkay.goralcan@pureenergy.com.tr, jane.sezgin@pureenergy.com.tr, tr-intraday@pureenergy.com.tr, can.ulutas@pureenergy.com.tr, umut.turabik@pureenergy.com.tr

to_emails = ["tr-intraday@pureenergy.com.tr", "operasyon@pureenergy.com.tr", "originationtr@pureenergy.com.tr"]
cc_emails = ["yaman.bicer@puretechnology.com.tr", "fatih.gultekin@pureenergy.com.tr", "feyza.kesilmis@pureenergy.com.tr"]

to_emails = ["yaman.bicer@puretechnology.com.tr", "fatih.gultekin@pureenergy.com.tr", "feyza.kesilmis@pureenergy.com.tr"]
cc_emails = ["yamanbicer@gmail.com"]
Cc: feyza.kesilmis@pureenergy.com.tr
"Gültekin, Fatih (PURE)" <fatih.gultekin@pureenergy.com.tr>
Mail text: 
Dear Puritans,

Attached is the Türkiye DSG P&L Daily Report for [Date]. It provides an overview of the previous day's management performance, including the daily P&L amount.

Best,
Pure Energy BI Team