JavaScript kann über ActiveXObject(...) auf einen ziemlich großen Teil der klassischen Windows-COM-Welt zugreifen. ADO ist dabei nur eine dieser COM-Bibliotheken.

Wichtig ist die Trennung:

HTA / JavaScript
      │
      └── ActiveXObject("...")
               │
               └── COM / ActiveX
                      ├── ADO → Datenbanken
                      ├── Excel
                      ├── Word
                      ├── Access
                      ├── Outlook
                      ├── FileSystemObject
                      ├── WScript.Shell
                      ├── Shell.Application
                      └── weitere installierte COM-Komponenten
1. ADO – Datenbanken

Das, was wir gerade benutzt haben:

var conn = new ActiveXObject("ADODB.Connection");
var rs   = new ActiveXObject("ADODB.Recordset");

Damit kannst du grundsätzlich alles ansprechen, wofür ein passender OLE-DB- oder ODBC-Treiber installiert ist.

Also beispielsweise:

SQLite       ← über SQLite ODBC-Treiber
Access       ← ACE / Jet
SQL Server
MySQL        ← ODBC-Treiber
PostgreSQL   ← ODBC-Treiber
Excel-Datei  ← kann sogar als Datenquelle behandelt werden
CSV/Text     ← ebenfalls möglich

Das Interessante ist also:

ADO ist nicht SQLite. ADO ist die allgemeine Datenbank-Schnittstelle.

SQLite hängt nur als Treiber dahinter.

2. Microsoft Excel

Direkt die komplette Excel-Automation:

var excel = new ActiveXObject("Excel.Application");


excel.Visible = true;


var wb = excel.Workbooks.Add();
var sheet = wb.Worksheets(1);


sheet.Cells(1, 1).Value = "Hallo aus JavaScript";
sheet.Cells(2, 1).Value = 42;

Das ist genau die Welt, die dir so VBA-artig vorkam.

JavaScript steuert hier dieselben COM-Objekte, die VBA innerhalb von Excel benutzt:

Application
   ↓
Workbook
   ↓
Worksheet
   ↓
Range / Cells

Darum fühlt sich das so ähnlich an.

3. Word

Dasselbe mit Word:

var word = new ActiveXObject("Word.Application");


word.Visible = true;


var doc = word.Documents.Add();
doc.Content.Text = "Hallo aus meiner HTA!";

Damit könntest du beispielsweise aus deiner HTA automatisch Berichte erzeugen.

4. Access

Auch Access selbst lässt sich automatisieren:

var access = new ActiveXObject("Access.Application");


access.Visible = true;
access.OpenCurrentDatabase("C:\\daten\\test.accdb");

Wobei du Access-Datenbanken häufig gar nicht über die Access-Anwendung öffnen musst.

Stattdessen:

HTA
 ↓
JavaScript
 ↓
ADO
 ↓
ACE
 ↓
.accdb

Das ist für reine Datenbankarbeit meist sauberer.

5. Windows-Dateisystem

Für HTAs extrem praktisch:

var fso = new ActiveXObject("Scripting.FileSystemObject");

Dann etwa:

if (fso.FileExists("C:\\test.txt")) {
    alert("Datei existiert");
}

Oder Dateien schreiben:

var file = fso.CreateTextFile("C:\\test.txt", true);
file.WriteLine("Hallo Welt");
file.Close();

Das ist ein gewaltiger Unterschied zwischen einer normalen Webseite und HTA.

Eine Webseite darf das nicht einfach.

Eine HTA schon.

6. Windows Shell

Ebenfalls sehr mächtig:

var shell = new ActiveXObject("WScript.Shell");

Damit kannst du Programme starten:

shell.Run("notepad.exe");

Oder:

shell.Run("cmd.exe");

Auch Umgebungsvariablen und Registry-Zugriffe gehören zu dieser COM-Welt.

7. Explorer / Windows Shell

Noch eine andere Schnittstelle:

var shell = new ActiveXObject("Shell.Application");

Damit kommst du an Windows-Shell-Funktionen, Ordner, Explorer usw.

8. Outlook

Wenn klassisches Outlook mit entsprechender COM-Automation installiert ist:

var outlook = new ActiveXObject("Outlook.Application");

Dann kannst du z. B. Mail-Objekte erzeugen und andere Outlook-Objekte automatisieren. Welche Office-COM-Schnittstellen tatsächlich verfügbar sind, hängt von der installierten Office-Version ab.

Und jetzt kommt der entscheidende Punkt

ActiveXObject() ist eigentlich eine Brücke aus JavaScript in eine komplett andere Objektwelt.

Wenn du schreibst:

var conn = new ActiveXObject("ADODB.Connection");

ist conn kein normales JavaScript-Objekt.

Bei:

var excel = new ActiveXObject("Excel.Application");

ist excel ebenfalls kein JavaScript-Objekt.

JavaScript hält jeweils nur eine Referenz auf ein COM-Objekt.

Deshalb sehen Sachen wie:

rs.Fields("name").Value
rs.MoveNext()
rs.EOF


excel.Workbooks.Add()
sheet.Cells(1,1).Value

so verdächtig nach VBA aus.

Das IST dieselbe COM-Objektwelt.

Nur der Aufrufer ist ein anderer:

VBA ───────────────┐
VBScript ──────────┤
JavaScript in HTA ─┼──→ COM
C++ ───────────────┤
PowerShell ────────┘

Und genau deshalb würde ich bei deinem HTA-Lernen die Grenze bewusst so betrachten:

// JavaScript-Welt
var name = "Maik";
var users = [];
var html = "";




// ----- COM-Grenze -----


var conn = new ActiveXObject("ADODB.Connection");
conn.Open(...);


var rs = conn.Execute("SELECT * FROM users");




// ----- zurück in JavaScript-Welt -----


while (!rs.EOF) {
    users.push({
        id: rs.Fields("id").Value,
        name: rs.Fields("name").Value
    });


    rs.MoveNext();
}

Das ist ein sehr gutes Muster: COM möglichst nur als Ein-/Ausgangsschicht benutzen und die Werte schnell in normale JS-Objekte/Arrays umwandeln.

Dann bleibt deine Anwendung überwiegend JavaScript und nur ein kleiner Rand ist dieses etwas schrullige VBA/COM-Land. 😄

zeigmirnochmal wie ich sqlite nutze

Klar. Für HTA + JavaScript + SQLite läuft es typischerweise so:

HTA / JavaScript
    ↓
ADODB.Connection
    ↓
ODBC-Treiber für SQLite
    ↓
deine .db-Datei

Minimal:

<hta:application>
<meta http-equiv="X-UA-Compatible" content="IE=9">


<script>
var conn = new ActiveXObject("ADODB.Connection");


conn.Open(
    "Driver=SQLite3 ODBC Driver;" +
    "Database=C:\\daten\\test.db;"
);


var rs = conn.Execute("SELECT * FROM users");


while (!rs.EOF) {
    alert(rs.Fields("name").Value);
    rs.MoveNext();
}


rs.Close();
conn.Close();
</script>

Der wichtigste Teil ist:

var conn = new ActiveXObject("ADODB.Connection");

Das erstellt die ADO-Verbindung. Dann:

conn.Open(
    "Driver=SQLite3 ODBC Driver;" +
    "Database=C:\\daten\\test.db;"
);

öffnet deine SQLite-Datenbank über den installierten SQLite-ODBC-Treiber.

Eine Abfrage machst du dann mit:

var rs = conn.Execute("SELECT * FROM users");

rs bedeutet Recordset. Das ist das Ergebnis deiner SQL-Abfrage.

Darüber iterierst du:

while (!rs.EOF) {
    var name = rs.Fields("name").Value;


    alert(name);


    rs.MoveNext();
}

Also gedanklich:

conn
 │
 └── Verbindung zur Datenbank


conn.Execute(...)
 │
 └── rs = Ergebniszeilen


rs.Fields("name").Value
 │
 └── Feld der aktuellen Zeile


rs.MoveNext()
 │
 └── nächste Zeile


rs.EOF
 │
 └── Ende erreicht

Für INSERT, UPDATE oder DELETE brauchst du gar kein Recordset:

conn.Execute(
    "INSERT INTO users (name) VALUES ('Maik')"
);

Sauberer würde ich es in deiner HTA ungefähr so kapseln:

var db = new ActiveXObject("ADODB.Connection");


function dbOpen() {
    db.Open(
        "Driver=SQLite3 ODBC Driver;" +
        "Database=C:\\daten\\test.db;"
    );
}


function dbClose() {
    db.Close();
}


function getUsers() {
    var rs = db.Execute("SELECT id, name FROM users");
    var users = [];


    while (!rs.EOF) {
        users.push({
            id: rs.Fields("id").Value,
            name: rs.Fields("name").Value
        });


        rs.MoveNext();
    }


    rs.Close();


    return users;
}
Dann ist der normale JavaScript-Teil angenehm:

dbOpen();


var users = getUsers();


for (var i = 0; i < users.length; i++) {
    alert(users[i].name);
    }


    dbClose();

    Das ist wahrscheinlich das Grundmuster, das du dir merken solltest:

    var conn = new ActiveXObject("ADODB.Connection");


    conn.Open("...");
    var rs = conn.Execute("SELECT ...");


    while (!rs.EOF) {
        var wert = rs.Fields("spalte").Value;
            rs.MoveNext();
            }


            rs.Close();
            conn.Close();

            Das reicht erstmal völlig. Alles andere ist Ausbau.
