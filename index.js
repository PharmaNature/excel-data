const xlsx = require('xlsx');

const fs = require('fs');
const path = require('path');
const nodemailer = require("nodemailer");
require('dotenv').config()
const AWS = require('aws-sdk');

AWS.config.update({
    accessKeyId: process.env.ACCESS_KEY_ID,
    secretAccessKey: process.env.SECRET_ACCESS_KEY,
    region: 'eu-west-3'
});

// temps début du process
const startTime = new Date()
console.log("Begin...");

// Chemin vers votre fichier Excel
const filePath = 'export.xlsx';

// Charger le fichier Excel
const workbook = xlsx.readFile(filePath);

// Obtenir le nom de la première feuille de calcul
const worksheet = workbook.Sheets[workbook.SheetNames[2]];

/**
 * FONCTIONS
 */

function getData() {
    console.log("Sheet to JSON");
    const dataBrut = xlsx.utils.sheet_to_json(worksheet);

    let data = []
    for (let i = 0; i < dataBrut.length; i++) {
        data[i] = {};
        let date = dataBrut[i]['Lignes de facture/Créé le']
        data[i]['Date'] = (new Date((date - 25569) * 86400 * 1000)).toISOString()
        data[i]['Produit'] = dataBrut[i]['Lignes de facture/Article']
        data[i]['Quantite'] = dataBrut[i]['Quantité']
        data[i]['Equipe'] = dataBrut[i]['Équipe']
        data[i]['Client'] = dataBrut[i]['ID client']
        data[i]['Prix'] = dataBrut[i]['Sous-total signé']

    }

    return data;
}

// Calcule temps 
// retourne la temps entre les deux dates
function getTimeProcess(startTime, endTime) {
    const elapsedTime = endTime - startTime;

    // Convertir la durée en millisecondes en une chaîne formatée pour l'afficher dans la console
    const formattedElapsedTime = new Date(elapsedTime).toISOString().substr(11, 8);
    console.log("Le processus à duré : " + formattedElapsedTime)
}
// Récupère toutes les année pr"sente dans le fichier excel 
//récupère toutes les années présente dans le fichier excel


function getYears() {
    const tab = [];
    const colonne = 'Date';

    for (let i = 0; i < data.length; i++) {
        const valeur = data[i][colonne];
        // Convertir la date en format YYYY-MM-DD en un tableau de chaînes de caractères
        const annee = valeur.substring(0, 4);;

        // Vérifier si la valeur n'est pas déjà présente dans le tableau
        if (!tab.includes(annee)) {
            tab.push(annee);
        }
    }
    tab.sort()
    return tab;
}

function getTeams() {
    const tab = [];
    const colonne = 'Equipe';

    for (let i = 0; i < data.length; i++) {
        const valeur = data[i][colonne];

        // Vérifier si la valeur n'est pas déjà présente dans le tableau
        if (!tab.includes(valeur) && valeur !== 'VNC' && valeur !== 'INTER') {
            tab.push(valeur);
        }
    }

    return tab;
}

function getProducts() {
    const tab = [];
    const colonne = 'Produit';

    for (let i = 0; i < data.length; i++) {
        const valeur = data[i][colonne];

        // Vérifier si la valeur n'est pas déjà présente dans le tableau
        if (!tab.includes(valeur)) {
            tab.push(valeur);
        }
    }

    return tab;
}

// top 25 des produits par années 
function getQtyPerYears() {
    let gigaDic = {}

    for (let k = 0; k < years.length; k++) {
    let microDic = {}
        
        for (let i = 0; i < data.length; i++) {
            let obj = data[i]
            let annee = obj['Date'].substring(0, 4)
            let produit = obj['Produit']
            let quantite = obj['Quantite']

            
            // si il n'y a pas l'année pour ce produit 
            if (!microDic[produit] && annee === years[k]) {
                microDic[produit] = 0
            }
            if (annee === years[k]){
            // ajoute la quantité pour la bonne année
            microDic[produit] += quantite
                
            }
            
        }
        gigaDic[years[k]] = microDic
    }
    return gigaDic
}

function getTop25PerYears() {
    let dic = getQtyPerYears()
    let dick = []
    for(let i = 0;i < years.length;i++){
        const products = dic[years[i]];
        const sortedProducts = Object.keys(products).sort((a, b) => products[b] - products[a]);
        dick[i] = sortedProducts.slice(0, 25);
    }
    return dick 
}

function getPricePerYears() {
    let gigaDic = {}

    for (let k = 0; k < years.length; k++) {
    let microDic = {}
        
        for (let i = 0; i < data.length; i++) {
            let obj = data[i]
            let annee = obj['Date'].substring(0, 4)
            let produit = obj['Produit']
            let quantite = obj['Quantite']
            let prix = obj['Prix']
            
            if (!isNaN(parseFloat(prix))) {
                prix = parseFloat(prix);
            } else {
                prix = 0;
            }
            
            // si il n'y a pas l'année pour ce produit 
            if (!microDic[produit] && annee === years[k]) {
                microDic[produit] = 0
            }
            if (annee === years[k]){
            // ajoute la quantité pour la bonne année
            microDic[produit] += prix
                
            }
            
        }
        gigaDic[years[k]] = microDic
    }
    return gigaDic
}

function getTop25PricePerYears(){
    let listePrix = getPricePerYears()
    let dic = {}
    for (let i = 0; i < years.length; i++) {
        dic[years[i]] = {}
        for (let j = 0; j < top25[0].length; j++) {
            dic[years[i]][top25[i][j]] = parseFloat((listePrix[years[i]][top25[i][j]]).toFixed(2))
        }
    }
    return dic

}


function getCustomers(team,year){
    let listeClient = []

    for (let i = 0; i < data.length; i++) {
        if (data[i]['Equipe'] === team && data[i]['Date'].substring(0,4) === year && !listeClient.includes(data[i]['Client'])) {
            listeClient.push(data[i]['Client'])
        }
    }

    return listeClient.length
}

function getAllCustomersMore(year) {
    let listeClient = []

    for (let i = 0; i < data.length; i++) {
        for (let j = 0; j < teams.length; j++) {
            if (data[i]['Equipe'] === teams[j] && data[i]['Date'].substring(0, 4) === year && !listeClient.includes(data[i]['Client'])) {
                listeClient.push(data[i]['Client'])
            }
        }
    }
    return listeClient
}

function getAllLostCustomers(){
    let dic = {}
    let dicLost = {}
    // pour chaque année tous les clients 
    for (let i = 0; i < years.length; i++) {
        dic[years[i]] = getAllCustomersMore(years[i])
        dicLost[years[i]] = []
    }

    // pour chaque année tous les clients qui ne sont plus là
    for (let i = 0; i < years.length-1; i++) {
        for (let j = 0; j < dic[years[i]].length; j++) {
            // vérifie pour tous les clients si ils sont dans l'année d'après
            if (!dic[years[i+1]].includes(dic[years[i]][j])) {
                dicLost[years[i+1]].push(dic[years[i]][j])
            }
        }
    }

    //console.log(dicLost);
    return dicLost
}

function getAllNewCustomers() {
    let dic = {};
    let dicNew = {};

    // pour chaque année tous les clients
    for (let i = 0; i < years.length; i++) {
        dic[years[i]] = getAllCustomersMore(years[i]);
        dicNew[years[i]] = [];
    }

    // pour chaque année tous les nouveaux clients
    for (let i = 0; i < years.length - 1; i++) {
        for (let j = 0; j < dic[years[i + 1]].length; j++) {
            // vérifie pour tous les clients s'ils sont nouveaux
            if (!dic[years[i]].includes(dic[years[i + 1]][j])) {
                dicNew[years[i + 1]].push(dic[years[i + 1]][j]);
            }
        }
    }

    //console.log(dicNew);
    return dicNew
}

function statsByTeamForNewAndLostCustomers() {
    let dicNew = getAllNewCustomers()
    let dicLost = getAllLostCustomers()

    let dicStats = {}
    // init dicStats
    for (let p = 1; p < years.length; p++) {
        dicStats[years[p]] = {}
        for (let t = 0; t < teams.length; t++) {
            dicStats[years[p]][teams[t]] = { lost:0, new:0, ratioLost:'', ratioNew:'', 'ratioLost%':'', 'ratioNew%':''}
        }
    }

    // NEW
    // parcours de toutes les données pour associer les nouveaux clients aux commerciaux 
    for (let p = 1; p < years.length; p++) {
        for (let x = 0; x < dicNew[years[p]].length; x++) {
            let notFound = true
            for (let i = 0; i < data.length && notFound; i++) {
                if (data[i]['Client'] == dicNew[years[p]][x]) {
                    dicStats[years[p]][data[i]['Equipe']]['new'] += 1
                    notFound = false
                }
            }
        }
    }

    // LOST
    // parcours de toutes les données pour associer les nouveaux clients aux commerciaux 
    for (let p = 1; p < years.length; p++) {
        for (let x = 0; x < dicLost[years[p]].length; x++) {
            let notFound = true
            for (let i = 0; i < data.length && notFound; i++) {
                if (data[i]['Client'] == dicLost[years[p]][x]) {
                    dicStats[years[p]][data[i]['Equipe']]['lost'] += 1
                    notFound = false
                }
            }
        }
    }

    // Ratio 
    for (let p = 1; p < years.length; p++) {
        for (let t = 0; t < teams.length; t++) {
            dicStats[years[p]][teams[t]]['ratioNew'] = dicStats[years[p]][teams[t]]['new'] + "/" + dicNew[years[p]].length
            dicStats[years[p]][teams[t]]['ratioLost'] = dicStats[years[p]][teams[t]]['lost'] + "/" + dicLost[years[p]].length
            dicStats[years[p]][teams[t]]['ratioNew%'] = ((dicStats[years[p]][teams[t]]['new']/dicNew[years[p]].length)*100).toFixed(1)+'%'
            dicStats[years[p]][teams[t]]['ratioLost%'] = ((dicStats[years[p]][teams[t]]['lost']/dicLost[years[p]].length)*100).toFixed(1) +'%'
        }
    }

    //console.log(dicStats);
    return dicStats
}

function getAllCustomers(){
    let dic = {}

    for (let i = 0; i < years.length; i++) {
        let year = years[i]

        dic[year] = {}
        for (let k = 0; k < teams.length; k++) {
            let team = teams[k]
            let nbCustomers = getCustomers(team,year)
            dic[year][team] = nbCustomers
            
        }
        
    }
    //console.log(dic);
    return dic
}

function getCustomersProduct(team,year,product){
    let listeClient = []

    for (let i = 0; i < data.length; i++) {
        count++
        if (data[i]['Equipe'] === team && data[i]['Date'].substring(0,4) === year && data[i]['Produit'] === product && !listeClient.includes(data[i]['Client'])) {
            listeClient.push(data[i]['Client'])
        }
    }

    return listeClient.length
}

function getAllCustomersProduct(){
    let dic = {}

    for (let i = 0; i < years.length; i++) {
        let year = years[i]
        let products = top25[i]
        dic[year] = {}
        for (let k = 0; k < teams.length; k++) {
            let team = teams[k]
            dic[year][team] = {}
            for (let j = 0; j < 25; j++) {
                
                let product = products[j]
                let nbCustomers = getCustomersProduct(team,year,product)
                dic[year][team][product] = nbCustomers
            }
            
            
        }
        
    }
    return dic;
}

function nationalStat(){
    let stat = {}
    stat['global'] = {} 
    stat['detail'] = {}

    for (let i = 0; i < years.length; i++) {
        stat['detail'][years[i]] = {}
        stat['global'][years[i]] = {}
        let sommeGlobal = 0

        for (let j = 0; j < teams.length; j++) {
                sommeGlobal += allCustomers[years[i]][teams[j]]

                for (let k = 0; k < top25[0].length; k++) {

                    if (!stat['detail'][years[i]][top25[i][k]]) {
                        stat['detail'][years[i]][top25[i][k]] = 0
                    }
                stat['detail'][years[i]][top25[i][k]] += allCustomersProduct[years[i]][teams[j]][top25[i][k]]
                }
        }
        stat['global'][years[i]] = sommeGlobal
    }
    return stat
}
// i = bloc d'équipe top 25 / j = top 25 produit en fonction de l'année /  k = nombre de colonne de top25 
function createDataTop25() {
    let excelRows = []
    ////////////////////////////////////////////////
    // TOP 25 des produits / année / équipe avec des stats de client 
    ////////////////////////////////////////////////
    // Pour chaque équipe 
    for (let i = 0; i < teams.length; i++) {
        // Ajouter une ligne vide pour séparer les blocs
        excelRows.push([]);
        let firstRow = []
        // Ajouter le nom de l'équipe
        firstRow.push([teams[i]]);
        for (let t = 0; t < years.length; t++) {
            firstRow.push('Produit', 'vendu/client total', 'vendu/client total %', '', '')
        }
        excelRows.push(firstRow)

        // Pour chaque année, afficher les produits du top 25 en colonnes
        for (let j = 0; j < top25[0].length; j++) {

            let row = [];
            // Ajouter les produits du top 25 en colonnes
            for (let k = 0; k < years.length; k++) {
                let ratio = allCustomersProduct[years[k]][teams[i]][top25[k][j]] + '/' + allCustomers[years[k]][teams[i]]
                let ratioPourcentage = allCustomersProduct[years[k]][teams[i]][top25[k][j]] / allCustomers[years[k]][teams[i]] * 100
                let datte = ''
                if (j === 0) {
                    datte = years[k]
                }
                row.push(datte, top25[k][j], ratio, ratioPourcentage.toFixed(0) + '%', '');
            }
            excelRows.push(row);
        }
    }

    ////////////////////////////////////////////////
    // TOP 25 des produits / année / stats national de client 
    ////////////////////////////////////////////////
    // Ajouter une ligne vide pour séparer les blocs
    excelRows.push([]);
    let firstRow = []
    // Ajouter le nom de l'équipe
    firstRow.push("NATIONAL ");
    for (let t = 0; t < years.length; t++) {
        firstRow.push('Produit', 'vendu/client total', 'vendu/client total %', 'Prix rapporté', '')
    }
    excelRows.push(firstRow)

    // Pour chaque année, afficher les produits du top 25 en colonnes

    let stats = nationalStat()
    for (let j = 0; j < top25[0].length; j++) {

        let row = [];
        // Ajouter les produits du top 25 en colonnes
        for (let k = 0; k < years.length; k++) {
            let ratio = stats['detail'][years[k]][top25[k][j]] + '/' + stats['global'][years[k]]
            let ratioPourcentage = stats['detail'][years[k]][top25[k][j]] / stats['global'][years[k]] * 100
            let datte = ' '
            if (j === 0) {
                datte = years[k]
            }
            row.push(datte, top25[k][j], ratio, ratioPourcentage.toFixed(0) + '%', top25Price[years[k]][top25[k][j]] + '€');
        }
        excelRows.push(row);

    }

    ////////////////////////////////////////////////
    // Stats sur les nouveaux clients et les clients perdus pour chaque équipe par année
    ////////////////////////////////////////////////
    // Ajouter une ligne vide pour séparer les blocs
    excelRows.push([]);
    excelRows.push([]);
    excelRows.push([]);
    let firstRowbis = []
    // Ajouter le nom de l'équipe
    firstRowbis.push("Clients Perdus/Gagnés ");
    for (let t = 0; t < years.length; t++) {
        firstRowbis.push('Equipe', 'Clients perdus', 'Clients perdus%', 'Clients gagnés','Clients gagnés%', '')
    }
    excelRows.push(firstRowbis)

    // Pour chaque année, afficher les produits du top 25 en colonnes

    let dicNewLost = statsByTeamForNewAndLostCustomers()
    for (let t = 0; t < teams.length; t++) {
        let row = []
        for (let k = 0; k < years.length; k++) {
            let dattte = ''
            if (k == 0) {
                if (t === 0) {
                    dattte = years[k]
                }
                row.push(dattte, '', '', '', '', '')
            }
            else {
                let cell = dicNewLost[years[k]][teams[t]]
                let datte = ' '
                if (t === 0) {
                    datte = years[k]
                }
                row.push(datte, teams[t], cell['ratioLost'], cell['ratioLost%'], cell['ratioNew'], cell['ratioNew%'])
            }

        }
        excelRows.push(row);
    }

    // pour avoir la liste de tous les clients partis et arrivés 
    excelRows.push([]);
    excelRows.push([]);
    let ligneUne = []
    // Ajouter le nom de l'équipe
    ligneUne.push("Clients Perdus/Gagnés détails ");
    for (let t = 0; t < years.length; t++) {
        ligneUne.push('', 'Clients perdus', '', 'Clients gagnés','', '')
    }
    excelRows.push(ligneUne)
    let dicNew = getAllNewCustomers()
    let dicLost = getAllLostCustomers()
    // avoir le max
    let max = 0
    for (let p = 1; p < years.length; p++) {
        if (dicNew[years[p]].length > max) {
            max = dicNew[years[p]].length 
        }
        if (dicLost[years[p]].length > max) {
            max = dicLost[years[p]].length 
        }
    }
    //console.log(max);
        for (let i = 0; i < max ; i++) {
            let row = []
            row.push('')
            for (let k = 0; k < years.length; k++) {
                let cellNew = ''
                let cellLost = ''
                if (i < dicNew[years[k]].length) {
                    cellNew = dicNew[years[k]][i]
                }if (i < dicLost[years[k]].length) {
                    cellLost = dicLost[years[k]][i]
                }
                //console.log(cellNew);
                //console.log(cellLost);
                row.push('',cellLost,'',cellNew,'','')
            }
            excelRows.push(row)
    }

    
    
    //console.log(excelRows);
    return excelRows;
}

// Génère un fichier excel avec le tableau qu'on lui donne
function generateOutput(data){
    // Créer un nouveau classeur
    const newWorkbook = xlsx.utils.book_new()
    console.log("generate Excel 1/5");
    // Créer une nouvelle feuille
    const newWorksheet = xlsx.utils.aoa_to_sheet(data)
    console.log("generate Excel 2/5");

    // Ajouter la nouvelle feuille au classeur 
    xlsx.utils.book_append_sheet(newWorkbook,newWorksheet, 'Feuille 1');
    console.log("generate Excel 3/5");

    // Chemin du fichier de sortie
    const rd = Math.floor(Math.random() * 10000);
    const outputFilePath = "./excelGenerated/top25_"+ rd +".xlsx"
    console.log("generate Excel 4/5    "+outputFilePath);

    // Enregistrer le classeur dans un fichier Excel 
    xlsx.writeFile(newWorkbook,outputFilePath)
    console.log("generate Excel 5/5");

    return outputFilePath
}


/**
 * CONSTANTES
 */
let count = 0
const data = getData()
console.log("Processing...");
const years = getYears().sort()
const teams = getTeams().sort()
const products = getProducts().sort()
const top25 = getTop25PerYears()
const allCustomers = getAllCustomers()
const allCustomersProduct = getAllCustomersProduct()
const top25Price = getTop25PricePerYears()

// TEST 



//statsByTeamForNewAndLostCustomers()

generateOutput(createDataTop25())

//console.log(count);
getTimeProcess(startTime, new Date())


AWS.config.update({
    accessKeyId: process.env.ACCESS_KEY_ID,
    secretAccessKey: process.env.SECRET_ACCESS_KEY,
    region: 'eu-west-3' // Remplacez par votre région AWS
});

const filePath2 = path.join(__dirname, './excelGenerated/');

fs.readdir(filePath2, (err, files) => {
    console.log(files);

    // Créer les paramètres pour l'envoi de l'e-mail avec la pièce jointe
    let usefulData = 'some,stuff,to,send';

    let transporter = nodemailer.createTransport({
        SES: new AWS.SES({ region: 'eu-west-3', apiVersion: "2010-12-01" })
    });

    let text = 'Attached is a CSV of some stuff.';

    // send mail with defined transport object
    transporter.sendMail({
        from: '"clement gras" <cgras@pharmanature.fr>',
        to: "cgras@pharmanature.fr",
        subject: "Hello",                // Subject line
        text: text,                      // plaintext version
        html: '<div>' + text + '</div>', // html version
        attachments: [{
            filename: files[0],
            content: fs.readFileSync(filePath2 + files[0])
        }]
    }).then(() => {
        // Supprimez le fichier Excel une fois que l'e-mail est envoyé
        fs.unlinkSync(filePath2 + files[0]);
    }).catch(error => {
        console.log(error);
    });
})

