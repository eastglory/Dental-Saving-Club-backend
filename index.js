const express = require('express')
const bodyParser = require('body-parser')
const cors = require('cors')
// const adodb = require('node-adodb')
const readXlsxFile = require('read-excel-file/node')
const XlsxPopulate = require('xlsx-populate')
const mysql = require('mysql')

const app = express()
const port = 4000;

app.use(cors())

// Configuring body parser middleware
app.use(bodyParser.urlencoded({ extended: true }))
app.use(bodyParser.json())

let con = mysql.createPool({
    host: "bvqj0bxlolitaru3lp3u-mysql.services.clever-cloud.com",
    user: "uni3mzp1n8uqnxfx",
    password: "qQYf148dBmURzPpQZgy6",
    database: "bvqj0bxlolitaru3lp3u"
})

const formatDate = (date) => {
    if(date) return new Date(date).toISOString().split('T')[0]
    else return ''
}
// con.connect(err => {
//     if (err) throw err
// })

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
            if(typeof row[1] == "number") {
                const tray = row[1]?row[1]:""
                const client = row[2]?row[2]:""
                const notes = row[4]?row[4]:""
                const receptionDate = row[5]?row[5].replace(/(..)\-(..)\-(....)/, "$2-$1-$3"):""
                const recId = receptionDate.replace(/(..)\-(..)\-(....)/, "$2-$1-$3") + tray
                const location = String(row[6]).toUpperCase()
                const status = String(row[7]).trim().toUpperCase()
                const followUp = row[8]?row[8]:""
                    data.push({
                        tray,
                        client,
                        notes,
                        recId,
                        receptionDate,
                        location,
                        status,
                        followUp
                })
            }
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
            ws.cell(`F${tray+3}`).value(receptionDate.replace(/(..)\-(..)\-(....)/, "$2-$1-$3"));
            ws.cell(`G${tray+3}`).value(location);
            ws.cell(`H${tray+3}`).value(status);
            ws.cell(`I${tray+3}`).value(followUp);
            return workbook.toFileAsync('files/repairtray.xlsx')
        })

    return getTrayData(req, res)
}

const setRepairLog = async (req, res) => {
    const data = req.body
    console.log(req.body)
    const dataSql = "INSERT INTO repairauthdata (recId, quoteOn, quotevia, followOn, followVia, status, asOf, reportedBy, instruction, comment, signature, signDate, servicedBy, servicedDate) VALUES ?"
    const dataValue = Object.values(data)
    dataValue.splice(10,1)
    const detailSql = "INSERT INTO repairauthdetail (recId, description, serial, invoice, dop, warranty, cost, authorised) VALUES ?"
    
    con.query(dataSql, [[dataValue]], (err, result) => {  
        if (err) throw err
        const detailValue = data.tableData.map(item => {
            const arr = Object.values(item)
            arr.unshift(data.recId)
            return arr
        })
        console.log(detailValue)
        con.query(detailSql, [detailValue], (err, result) => {
            if (err) throw err
            return res.send(result)
        })
    })
    
}

const setRepairJournal = async (req, res) => {
    const data = req.body
    const sql = "INSERT INTO repairjournal (recid, datRec, datHan, datRep, person, client, invoice, serial, product, warranty, subject, failurDesc, malfunctioned, defect, comment, check1, check2, check3, check4, check5, bearing, chuck, waterblockage, lubrification, feasability, resn, images) VALUES ?"
    const value = Object.values(data)
    
    con.query(sql, [[value]], (err, result) => {  
        if (err) throw err
            return res.send(result)
        })
}

const getServiceData = async (req, res) => {
    const client = req.query.client
    const sql = `SELECT * FROM repairjournal WHERE client = '${client}'`
    con.query(sql, (err, result) => {
        if(err) throw err
        return res.send(result)
    })
}

const getSerialsFromRecId = async (req, res) => {
    const recId = req.query.recId
    const sql = `SELECT * FROM repairauthdetail WHERE recId = '${recId}'`
    con.query(sql, (err, result) => {
        if(err) throw err
        return res.send(result)
    })
}

const getRepairTrackerData = async (req, res) => {
    const sql = `SELECT client, product, serial, datRec, waterblockage, lubrification, bearing, chuck, feasability, resn, (SELECT dop FROM repairauthdetail WHERE repairauthdetail.serial = repairjournal.serial) AS dop FROM repairjournal`
    con.query(sql, (err, result) => {
        if(err) throw err
        return res.send(result)
    })
}

const getDashboardData = async (req, res) => {
    const {products, year} = req.body

    let authData 
    let authSql = `SELECT * from repairauthdetail WHERE description IN (`
    authSql += "\'" + products[0] + "\'"
    products.forEach(product => {
        authSql += "," + "\'" + product + "\'"
    })
    authSql += ")"

    let trackerData
    let trackerSql = `SELECT client, product, serial, datRec, waterblockage, lubrification, bearing, chuck, feasability, resn, (SELECT dop FROM repairauthdetail WHERE repairauthdetail.serial = repairjournal.serial) AS dop FROM repairjournal WHERE YEAR(datRec)=${year} AND product IN (`
    trackerSql += "\'" + products[0] + "\'"
    products.forEach(product => {
        trackerSql += "," + "\'" + product + "\'"
    })
    trackerSql += ")"

    con.query(authSql, (err, result) => {
        if(err) throw err
        authData = result.filter(item => parseInt(item.recId.split('-')[2].slice(0,4)) == year)
        con.query(trackerSql, (err, resp) => {
            if(err) throw err
            return res.send({authData, trackerData: resp})
    })
    })

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
app.get('/getserialsfromrecid', getSerialsFromRecId)
app.get('/getservicedata', getServiceData)
app.get('/getrepairtrackerdata', getRepairTrackerData)
app.post('/getdashboarddata', getDashboardData)
app.post('/settray', setTrayData)
app.post('/setRepairLog', setRepairLog)
app.post('/setRepairJournal', setRepairJournal)

app.listen(port, () => console.log(`Hello world app listening on port ${port}!`))