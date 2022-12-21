const express = require('express')
const bodyParser = require('body-parser')
const cors = require('cors')
// const adodb = require('node-adodb')
const readXlsxFile = require('read-excel-file/node')
const XlsxPopulate = require('xlsx-populate')

const app = express()
const port = 4000;

app.use(cors())

// Configuring body parser middleware
app.use(bodyParser.urlencoded({ extended: true }))
app.use(bodyParser.json())

//configuring database connection
// const connection = adodb.open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=C:\\\DSC\\Program\\Admin.accdb;Persist Security Info=False;", true)

// const getCustomer = async (req, res) => {
//     const customerId = req.query.customer
//     try {
//         const users = await connection.query(`SELECT * from Customer WHERE ID=${customerId}`)
//         return res.send(users)
//     } catch (error) {
//         console.error(error)
//     }
// }

// const getPurchaseInfo = async (req, res) => {
//     const customerId = req.query.customer
//     try{
//         const purchase = await connection.query(`SELECT * from Purchase WHERE Customer=${customerId}`)
//         return res.send(purchase)
//     } catch (err) {
//         console.error(err)
//     }
// }

const getTrayData = async (req, res) => {
    let data = [] ;
    await readXlsxFile('files/repairtray.xlsx').then((rows) => {
        rows.forEach((row, index) => {
            if(typeof row[1] == "number") data.push({
                tray: row[1],
                client: row[2],
                recId: row[3],
                notes: row[4]?row[4]:"",
                receptionDate: row[5]?row[5].replace(/(..)\-(..)\-(....)/, "$2-$1-$3"):"",
                location: String(row[6]).toUpperCase(),
                status: String(row[7]).trim().toUpperCase(),
                followUp: row[8]?row[8]:""
            })
        })
    })
    return res.send(data)
}

const setTrayData = async (req, res) => {
    const { tray, client, recId, notes, receptionDate, location, status, followUp} = req.body

    XlsxPopulate.fromFileAsync('files/repairtray.xlsx')
        .then(workbook => {
            let ws = workbook.sheet(0);
            ws.cell(`C${tray+3}`).value(client);
            ws.cell(`D${tray+3}`).value(recId);
            ws.cell(`E${tray+3}`).value(notes);
            ws.cell(`F${tray+3}`).value(receptionDate);
            ws.cell(`G${tray+3}`).value(location);
            ws.cell(`H${tray+3}`).value(status);
            ws.cell(`I${tray+3}`).value(followUp);
            return workbook.toFileAsync('files/repairtray.xlsx')
        })

    return getTrayData(req, res)
    
    
}


const getAllClients = async (req, res) => {
    let clients = []
    await readXlsxFile('files/customerv2.xlsx').then((rows) => {
        const temp = rows.slice(4, -2)
        clients = temp.map((client, index) => {
            return {
                id: index,
                name: client[0],
                street: client[1],
                phone: client[2],
                city: client[3],
                province: client[4],
                postal: client[5],
                country: client[7]
            }
        })
    })
    return res.send(clients)
}


// app.get('/customer', getCustomer)
// app.get('/purchase', getPurchaseInfo)
app.get('/getalltray', getTrayData)
app.get('/getallclients', getAllClients)
app.post('/settray', setTrayData)

app.listen(port, () => console.log(`Hello world app listening on port ${port}!`))