const XLSX = require('xlsx')
const fs = require('fs')

// Get file names
const folderPath = 'Files'
const filesList = []

fs.readdirSync(folderPath).forEach(file => {
    filesList.push(file)
});
//
// Cell names


// Лист Маршруты
const cellWarehouse = 'F2'

//Лист Расширенные маршруты
const cellDate = 'K3'

// Лист Метрики
const cellCouriersCount = 'C2'
const cellOrdersCount = 'C3'
const cellSumRouteLength = 'C10'

let resultArray = []

// Sheets' names
let SimpleRoutes = []
let ExtendedRoutes = []
let Metrics = []


// This is in test branch
for (let i = 0; i < filesList.length; i++) {
    const singleFile = XLSX.readFile(`${folderPath}/${filesList[i]}`)
    // Getting Sheet's data from single file
    SimpleRoutes = singleFile.Sheets['Маршруты']
    ExtendedRoutes = singleFile.Sheets['Расширенные маршруты']
    Metrics = (singleFile.Sheets['Метрики']) /* XLSX.utils.sheet_to_json */
    SimpleRoutesJSON = XLSX.utils.sheet_to_json(SimpleRoutes)



    // Creating AoA
    const oneDayData = []

    // Warehouse column
    const warehouseData = SimpleRoutes[cellWarehouse].v

    // Date coulumn (first 10 symbols )
    const date = ExtendedRoutes[cellDate].v.slice(0, 10)

    // Couriers column
    const couriersCount = Metrics[cellCouriersCount].v

    // Addresses column
    const addressArray = []
    SimpleRoutesJSON.forEach((address) => {
        if (address['Адрес']) {
            addressArray.push(address['Адрес'])
        }

    })
    const uniqueAddresses = [...new Set(addressArray)].length

    // Orders column
    const ordersCount = Metrics[cellOrdersCount].v

    // Lots count column
    const lotArray = []
    SimpleRoutesJSON.forEach((lotName) => {
        if (lotName['Комментарий']) {
            lotArray.push(lotName['Комментарий'])
        }

    })
    const uniqueLots = [...new Set(lotArray)].length


    // Travel distance culumn
    const traverDistanceNumber = parseInt((Metrics[cellSumRouteLength].v).replace(/\s/g, ''), 10)

    const travelDistanceAvergae = (traverDistanceNumber / couriersCount).toFixed(1)


    oneDayData.push(warehouseData, date, couriersCount, uniqueAddresses, ordersCount, uniqueLots, travelDistanceAvergae)

    // Pushing one day data as an array
    resultArray.push(oneDayData)

}
resultArray.unshift(['Склад', 'Дата', 'Количество курьеров', 'Количество точек', 'Количество заказов', 'Количество лотов', 'Пробег среднее'])


const newWS_name = 'Results'
const newWS = XLSX.utils.aoa_to_sheet(resultArray, { wpx: 200 })
newWS['!cols'] = [{ wpx: 85 }, { wpx: 80 }, { wpx: 120 }, { wpx: 100 }, { wpx: 110 }, { wpx: 100 }, { wpx: 90 }]
const newWB = XLSX.utils.book_new()

XLSX.utils.book_append_sheet(newWB, newWS, newWS_name)
XLSX.writeFile(newWB, './Result.xlsx')
