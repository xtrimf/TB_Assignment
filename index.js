const GoogleSpreadsheet = require("google-spreadsheet") // GOOGLE Sheets API
const {promisify} = require("util") // for async /await
const creds = require("./client_secret.json") // credential files
const fs = require('fs');

// prepare regex modules
const macRegExp = new RegExp(/(?:[a-z,A-Z,0-9]{4}.[a-z,A-Z,0-9]{4}?.[a-z,A-Z,0-9]{4}|([a-z,A-Z,0-9]{2}:){5}[a-z,A-Z,0-9]{2})/) //mac address validation in form of "xxxx.xxxx.xxxx" OR "xx:xx:xx:xx:xx:xx" - no checking for Hexadecimal chars
const ipRegExp = new RegExp(/\d{1,3}\.\d{1,3}\.\d{1,3}\.[0-9Z]{1,3}/) // ip address validation - exception for "Z" chars in 4th oct
const serialRegExp = new RegExp(/[A-Z,0-9]{7,8}/) //serial number valiation capital letters and/or numbers, allow length of 7-8 chars

let allRows = []


async function accessSpreadsheet() {
    
    //const doc = new GoogleSpreadsheet("13ITAyM-8pq9hay_Y2NNOqS_ps2rSlWR080yBEEviXmI") // public doc
    const doc = new GoogleSpreadsheet("1WaSeM0q_ecSqASt28sJvCUhyEfSeC9INRtZD4HBkwq4") // private doc

    // authenticate
    try {
        await promisify(doc.useServiceAccountAuth)(creds)
    } catch (e) {throw e};

    //get document info
    const info = await promisify(doc.getInfo)()

    //get and process sheets data and
    const headers = ["Electricity", "Name", "IPMI", "MAC", "Serial"];
    for (const sheet of info.worksheets) {
        let headerRow, h1, h2
        let rows = []
        let cellList = []
        // get cells data
        const cells = await promisify(sheet.getCells)({})

        cells.forEach(cell => {
            if (cell._value == headers[0]) {
                if (!h1) { // mark first "half"
                    h1 = cell.col
                    h2 = h1 + 5
                    lastCol = h2 + 4
                    headerRow = cell.row
                }
            }

            if (h1 && headerRow < cell.row && h1 <= cell.col && lastCol >= cell.col && cell._value != "empty") {
                cellList.push({
                    row: cell.row,
                    col: cell.col,
                    value: cell._value,
                    header: cell.col < h2 ? headers[cell.col - h1] : headers[cell.col - h2],
                    h: cell.col < h2 ? "h1" : "h2"
                });
            }
        });

        function template(data) {
            if (data.search(/SW - .{4}\..{4}\..{4}(?!.*TAG)/) > -1) return 1 // pattern #1 - SW with negative outlook of tags
            if (data.search(/\n\d{1,3}\.\d{1,3}\.\d{1,3}\.[0-9Z]{1,3}/) > -1) return 2 // pattern #2 - multiline
            if (data.search(/PDU [A-B] - (BLUE|RED)/) > -1) return 3 // pattern #3 - PDU with details
            if (data.search(/SW.+?TAG:/) > -1) return 4 // pattern #4 - SW with tags
            return 0 // no pattern - single cell distribution
        }

        function matchCell(data) {
            if (ipRegExp.exec(data)) return "ip"
            if (macRegExp.exec(data)) return "MAC"
            return 0
        }

        // start constructing the rows
        let row, col
        let i = 0
        while (i < cellList.length) {
            cellList[i].value
            let result = template(cellList[i].value);
            //console.log( cellList[i] , result);

            switch (result) {
                case 1:
                    row = cellList[i].value.split(" - ")
                    rows.push({
                        name: row[0],
                        ipmi: row[2],
                        MAC: row[1],
                        serial: row[3],
                        location: sheet.title
                    })
                    break;
                case 2:
                    row = cellList[i].value.split("\n")
                    rows.push({
                        name: row[0],
                        ipmi: ipRegExp.exec(cellList[i].value)[0],
                        MAC: macRegExp.exec(cellList[i].value)[0],
                        serial: cellList[i + 1].value,
                        location: sheet.title
                    })
                    i++ //advance one cell
                    break;
                case 3:
                    row = cellList[i].value.split(" - ")
                    rows.push({
                        name: cellList[i].value.match(/PDU [A-B] - (BLUE|RED)/)[0],
                        ipmi: ipRegExp.exec(cellList[i].value) != null ? ipRegExp.exec(cellList[i].value)[0] : null,
                        MAC: macRegExp.exec(cellList[i].value) != null ? macRegExp.exec(cellList[i].value)[0] : null,
                        serial: null,
                        location: sheet.title
                    })
                    break;
                case 4:
                    rows.push({
                        name: cellList[i].value.match(/\b.*SW/)[0].trim(),
                        ipmi: ipRegExp.exec(cellList[i].value) != null ? ipRegExp.exec(cellList[i].value)[0] : null,
                        MAC: macRegExp.exec(cellList[i].value) != null ? macRegExp.exec(cellList[i].value)[0] : null,
                        //serial: serialRegExp.exec(cellList[i].value) != null ? macRegExp.exec(cellList[i].value)[0] : null,
                        location: sheet.title
                    })
                    break;

                default:
                    let name = cellList[i].value.trim()
                    let ip = mac = serial = null
                    // match the rest of the cells after "name"
                    while (cellList[i + 1] && cellList[i].h == cellList[i + 1].h && cellList[i].row == cellList[i + 1].row) {
                        i++
                        ip = !ip && matchCell(cellList[i].value) == "ip" ? cellList[i].value.trim() : ip
                        mac = !mac && matchCell(cellList[i].value) == "MAC" ? cellList[i].value.trim() : mac
                        serial = !serial && matchCell(cellList[i].value) == 0 ? cellList[i].value.trim() : serial
                    }
                    rows.push({
                        name: name,
                        ipmi: ip,
                        MAC: mac,
                        serial: serial,
                        location: sheet.title
                    })
                    break;
            }
            i++
        }
        allRows = [...allRows, ...rows];
        console.log(`loading ${sheet.title}`)
        console.log(`${sheet.title} -> ${rows.length} Rows`);
    }
    fs.writeFile("results.json", JSON.stringify(allRows).replace(/},/g, '},\n'), function (err) {
        if (err) {
            return console.log(err);
        }
    });

}


// create the final SQL Statement
function createSQL() {
    let values = `CREATE TABLE IF NOT EXISTS t1 (
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        name VARCHAR(32),
        ipmi VARCHAR(32),
        MAC VARCHAR(32),
        serial VARCHAR(32),
        location VARCHAR(32) NOT NULL
    )  ENGINE=INNODB;
    
    INSERT INTO t1 (name,ipmi,MAC,serial,location) VALUES`

    for (let i = 0; i < allRows.length; i++) {
        let name = allRows[i].name ? `'${allRows[i].name}'` : null
        let ipmi = allRows[i].ipmi ? `'${allRows[i].ipmi}'` : null
        let MAC = allRows[i].MAC ? `'${allRows[i].MAC}'` : null
        let serial = allRows[i].serial ? `'${allRows[i].serial}'` : null

        values = `${values}(${name},${ipmi},${MAC},${serial},'${allRows[i].location}')${i < allRows.length - 1 ? ',\n' : ''}`

    }
    // write final SQL to file
    fs.writeFile("SQL_Insert.json", values, function (err) {
        if (err) {
            return console.log(err);
        } else {
            console.log('Done!')
        }
    });

    // code here to execute SQL...
}

//console.log(isIP("10.0.0.100"));
async function main() {
    await accessSpreadsheet()
    console.log('Generating SQL')
    createSQL()
}

main()
