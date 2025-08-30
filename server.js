const express = require('express');
const bodyParser = require('body-parser');
const XLSX = require('xlsx');
const cors = require('cors');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3000;

// Ordner für Excel-Datei
const excelFile = path.join(__dirname, 'familien.xlsx');

app.use(cors());
app.use(bodyParser.json());

// POST-Route zum Speichern der Daten
app.post('/save', (req, res) => {
    const { data } = req.body;

    let wb;
    if(fs.existsSync(excelFile)){
        wb = XLSX.readFile(excelFile);
    } else {
        wb = XLSX.utils.book_new();
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, 'Familien_' + Date.now());
    XLSX.writeFile(wb, excelFile);

    res.json({ message: 'Daten erfolgreich in Excel gespeichert!' });
});

app.listen(PORT, () => {
    console.log(`Server läuft auf http://localhost:${PORT}`);
});
