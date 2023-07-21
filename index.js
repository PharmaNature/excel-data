const xlsx = require('xlsx');
const XlsxPopulate = require('xlsx-populate');

// temps début du process
const startTime = new Date()
console.log("Begin...");






/**
 * FONCTIONS
 */

function getData() {
    
    // Chemin vers votre fichier Excel
const filePath = 'export_test_1.xlsx';

// Charger le fichier Excel
const workbook = xlsx.readFile(filePath);

// Obtenir le nom de la première feuille de calcul
const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    console.log("Sheet to JSON");
    const dataBrut = xlsx.utils.sheet_to_json(worksheet);
    let data = []
    const equipeBAN = ['VNC','INTER','COFFRES','FACONNAGE','MDD','PARTICULIERS','PERSONNEL', 'CD03']

    for (let i = 0; i < dataBrut.length; i++) {
        let date = dataBrut[i]['Lignes de facture/Créé le']
        let produit =  dataBrut[i]['Lignes de facture/Article']
        let quantite = dataBrut[i]['Lignes de facture/Quantité']
        let equipe = dataBrut[i]['Lignes de facture/Partenaire/Équipe commerciale']
        let client = dataBrut[i]['Lignes de facture/Partenaire/Référence']
        let prix = dataBrut[i]['Lignes de facture/Sous-total signé']
        let departement= dataBrut[i]['Lignes de facture/Partenaire/Code postal']
        if (typeof departement !== 'undefined' && typeof equipe !== 'undefined' && !equipeBAN.includes(equipe)) {
            data[i] = {};
            data[i]['Date'] = (new Date((date - 25569) * 86400 * 1000)).toISOString()
            data[i]['Produit'] = produit
            data[i]['Quantite'] = quantite
            data[i]['Equipe'] = equipe
            data[i]['Client'] = client
            data[i]['Prix'] = prix
            departement = departement.toString()
            if (departement.includes('AD')) {
                departement = '99'
            }
            else if (departement.length == 4) {
                departement = '0'+ departement
            }
            data[i]['Dpt'] = departement.substring(0,2)
        }
        
    }
    return data
}

function getDataEquipeDpt() {

    // Chemin vers votre fichier Excel
    const filePath = 'export_equipeDPT.xlsx';

    // Charger le fichier Excel
    const workbook = xlsx.readFile(filePath);

    // Obtenir le nom de la première feuille de calcul
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    console.log("Sheet to JSON ");
    const dataBrut = xlsx.utils.sheet_to_json(worksheet);

    //console.log(dataBrut);
    let data = []
    for (let i = 0; i < dataBrut.length; i++) {
        data[i] = {};
        data[i]['Secteur'] = dataBrut[i]['Secteur']
        data[i]['DPT'] = dataBrut[i]['Dépt.']
        data[i]['Potentiel'] = dataBrut[i]['Potentiel']
        data[i]['Actifs'] = dataBrut[i]['Actifs à fin 2022']
        data[i]['CA'] = dataBrut[i][' CAht 2022 ']
        data[i]['Couverture'] = dataBrut[i]['couverture']

    }

    return data;
}

function getEquipeDpt() {
    const dic = {};

    for (let i = 0; i < dataDPT.length; i++) {
        const secteur = dataDPT[i]['Secteur'];
        let dpt = dataDPT[i]['DPT'];

        if (secteur && dpt) {
            dpt = dpt.toString()
            if (dpt.length == 1) {
                dpt = "0"+dpt 
            }
            if (!dic[secteur]) {
                dic[secteur] = [];
            }
            dic[secteur].push(dpt);
        }
    }

    return dic;
}

// Calcule temps 
// retourne la temps entre les deux dates
function getTimeProcess(startTime, endTime) {
    const elapsedTime = endTime - startTime;

    // Convertir la durée en millisecondes en une chaîne formatée pour l'afficher dans la console
    const formattedElapsedTime = new Date(elapsedTime).toISOString().substr(11, 8);
    console.log("Le processus à duré : " + formattedElapsedTime)
}
// Récupère toutes les année présente dans le fichier excel 
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
    const equipeBAN = ['VNC','INTER','COFFRES','FACONNAGE','MDD','PARTICULIERS','PERSONNEL']
    for (let i = 0; i < data.length; i++) {
        if (data[i]) {
            const valeur = data[i][colonne];

            // Vérifier si la valeur n'est pas déjà présente dans le tableau
            if (!tab.includes(valeur) && !equipeBAN.includes(valeur)) {
                tab.push(valeur);
            }
        }
    }

    return tab;
}

function getDpts() {
    const tab = [];
    const colonne = 'Dpt';

    for (let i = 0; i < data.length; i++) {
        if (data[i]) {
            const valeur = data[i][colonne];

            // Vérifier si la valeur n'est pas déjà présente dans le tableau
            if (!tab.includes(valeur) && valeur) {
                tab.push(valeur);
            }
        }
    }

    return tab;
}

function getPotentielDpts(){
    const dic = {}

    for (let i = 0; i < dataDPT.length; i++) {
        if(dataDPT[i]['DPT'] && dataDPT[i]['Potentiel']){
            let departement = dataDPT[i]['DPT']
            if (departement.toString().length === 1) {
                departement = '0'+ departement
            }
            dic[departement] = dataDPT[i]['Potentiel']   
        }
    }

    return dic
}

function getCADpts(){
    const dic = {}
    for (let i = 0;i < dataDPT.length;i++){
        if(dataDPT[i]['DPT'] && dataDPT[i]['CA']){
            let departement = dataDPT[i]['DPT']
            dic[departement] = dataDPT[i]['CA']   
        }
    }
    return dic
}

function getProducts() {
    const tab = [];
    const colonne = 'Produit';

    for (let i = 0; i < data.length; i++) {
        if (data[i]) {
            const valeur = data[i][colonne];
            // Vérifier si la valeur n'est pas déjà présente dans le tableau
            if (!tab.includes(valeur)) {
                tab.push(valeur);
            }
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
function getTopPerYears() {
    let dic = getQtyPerYears()
    let dick = []
    for(let i = 0;i < years.length;i++){
        const products = dic[years[i]];
        const sortedProducts = Object.keys(products).sort((a, b) => products[b] - products[a]);
        dick[i] = sortedProducts
    }
    return dick 
}

function getPricePerDtp() {
    let gigaDic = {}

    for (let k = 0; k < years.length; k++) {
    let microDic = {}
        
        for (let i = 0; i < data.length; i++) {
            let obj = data[i]
            let annee = obj['Date'].substring(0, 4)
            let prix = obj['Prix']
            let dpt = obj['Dpt']
            
            if (!isNaN(parseFloat(prix))) {
                prix = parseFloat(prix);
            } else {
                prix = 0;
            }
            
            // si il n'y a pas l'année pour ce produit 
            if (!microDic[dpt] && annee === years[k]) {
                microDic[dpt] = 0
            }
            if (annee === years[k]){
            // ajoute la quantité pour la bonne année
            microDic[dpt] += parseInt(prix)
            }
            
        }
        gigaDic[years[k]] = microDic
    }
    
    return gigaDic
}

function getPricePerYears() {
    let gigaDic = {}

    for (let k = 0; k < years.length; k++) {
    let microDic = {}
        
        for (let i = 0; i < data.length; i++) {
            let obj = data[i]
            let annee = obj['Date'].substring(0, 4)
            let produit = obj['Produit']
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
function getTopAllPricePerYears(){
    let listePrix = getPricePerYears()
    let dic = {}
    for (let i = 0; i < years.length; i++) {
        dic[years[i]] = {}
        for (let j = 0; j < topAll[0].length; j++) {
            dic[years[i]][topAll[i][j]] = parseFloat((listePrix[years[i]][topAll[i][j]]).toFixed(2))
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

function getCustomersDpt(dpt, year) {
    let listeClient = []

    for (let i = 0; i < data.length; i++) {
        if (data[i]['Dpt']) {
            if (data[i]['Dpt'].toString().substring(0, 2) === dpt && data[i]['Date'].substring(0, 4) === year && !listeClient.includes(data[i]['Client'])) {
                listeClient.push(data[i]['Client'])
            }
        }
    }
    //console.log(listeClient);
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

function getAllCustomersDpt(){
    let dic = {}

    for (let i = 0; i < years.length; i++) {
        let year = years[i]

        dic[year] = {}
        for (let k = 0; k < dpts.length; k++) {
            let dpt = dpts[k]
            let nbCustomers = getCustomersDpt(dpt,year)
            dic[year][dpt] = nbCustomers
            
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
        let products = topAll[i]
        dic[year] = {}
        for (let k = 0; k < teams.length; k++) {
            let team = teams[k]
            dic[year][team] = {}
            for (let j = 0; j < products.length; j++) {
                
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

                for (let k = 0; k < topAll[0].length; k++) {

                    if (!stat['detail'][years[i]][topAll[i][k]]) {
                        stat['detail'][years[i]][topAll[i][k]] = 0
                    }
                stat['detail'][years[i]][topAll[i][k]] += allCustomersProduct[years[i]][teams[j]][topAll[i][k]]
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
    for (let j = 0; j < topAll[0].length; j++) {

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


function createDataDN() {
    let dataBySheet =  []
    
    ////////////////////////////////////////////////
    // Tous les produits / année / équipe avec des stats de client 
    ////////////////////////////////////////////////
    // Pour chaque équipe 
    for (let i = 0; i < teams.length; i++) {
    
        let excelRows = []
        // Ajouter une ligne vide pour séparer les blocs
        //excelRows.push([]);
        let firstRow = []
        // Ajouter le nom de l'équipe
        firstRow.push([teams[i]]);
        for (let t = 0; t < years.length; t++) {
            firstRow.push('Produit', 'vendu/client total', 'vendu/client total %', '', '')
        }
        excelRows.push(firstRow)

        // Pour chaque année, afficher les produits du top 25 en colonnes
        for (let j = 0; j < topAll[0].length; j++) {

            let row = [];
            // Ajouter les produits du top 25 en colonnes
            for (let k = 0; k < years.length; k++) {
                let ratio = allCustomersProduct[years[k]][teams[i]][topAll[k][j]] + '/' + allCustomers[years[k]][teams[i]]
                let ratioPourcentage = (allCustomersProduct[years[k]][teams[i]][topAll[k][j]] / allCustomers[years[k]][teams[i]] * 100).toFixed(1)
                let datte = ''
                if (j === 0) {
                    datte = years[k]
                }
                row.push(datte, topAll[k][j], ratio, ratioPourcentage + '%', '');
                if (years[k] && teams[i] && topAll[k][j] && ratioPourcentage) {
                dataForColor[years[k]][teams[i]][topAll[k][j]]  =  ratioPourcentage
                }
            }
            excelRows.push(row);
        }

        dataBySheet.push(excelRows)
    }
    
    ////////////////////////////////////////////////
    // Tous les produits / année / stats national de client 
    ////////////////////////////////////////////////
    // Ajouter une ligne vide pour séparer les blocs
    let excelRowsNational = []

    let firstRow = []
    // Ajouter le nom de l'équipe
    firstRow.push("NATIONAL");
    for (let t = 0; t < years.length; t++) {
        firstRow.push('Produit', 'vendu/client total', 'vendu/client total %', 'Prix rapporté', '')
    }
    excelRowsNational.push(firstRow)

    // Pour chaque année, afficher les produits du top All en colonnes

    let stats = nationalStat()
    //console.log(topAll);
    for (let j = 0; j < topAll[0].length; j++) {

        let row = [];
        // Ajouter les produits du top 25 en colonnes
        for (let k = 0; k < years.length; k++) {
            let ratio = stats['detail'][years[k]][topAll[k][j]] + '/' + stats['global'][years[k]]
            let ratioPourcentage = (stats['detail'][years[k]][topAll[k][j]] / stats['global'][years[k]] * 100).toFixed(1)
            let datte = ' '
            if (j === 0) {
                datte = years[k]
            }
            row.push(datte, topAll[k][j], ratio, ratioPourcentage + '%', topAllPrice[years[k]][topAll[k][j]] + '€');
            if (years[k] && topAll[k][j] && ratioPourcentage) {
                dataForColor[years[k]]["NATIONAL"][topAll[k][j]] =  ratioPourcentage                
            }
            
        }
        excelRowsNational.push(row);

    }

    dataBySheet.push(excelRowsNational)
    
    ////////////////////////////////////////////////
    // Stats sur les nouveaux clients et les clients perdus pour chaque équipe par année
    ////////////////////////////////////////////////
    // Ajouter une ligne vide pour séparer les blocs
    let excelRowsNewLost = []
    let firstRowbis = []
    // Ajouter le nom de l'équipe
    firstRowbis.push("Clients Perdus/Gagnés ");
    for (let t = 0; t < years.length; t++) {
        firstRowbis.push('Equipe', 'Clients perdus', 'Clients perdus%', 'Clients gagnés','Clients gagnés%', '')
    }
    excelRowsNewLost.push(firstRowbis)

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
        excelRowsNewLost.push(row);
    }

    // pour avoir la liste de tous les clients partis et arrivés 
    excelRowsNewLost.push([]);
    excelRowsNewLost.push([]);
    let ligneUne = []
    // Ajouter le nom de l'équipe
    ligneUne.push("Clients Perdus/Gagnés détails ");
    for (let t = 0; t < years.length; t++) {
        ligneUne.push('', 'Clients perdus', '', 'Clients gagnés','', '')
    }
    excelRowsNewLost.push(ligneUne)
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
            excelRowsNewLost.push(row)
    }
    dataBySheet.push(excelRowsNewLost)


    /////////////////////////////////////////
    // Département client et potentiel 
    /////////////////////////////////////////
    let excelRowsDpt = []

    // max key dpt 
    let maxDpt = 0
    for (let p = 0; p < years.length; p++) {
        let long = Object.keys(allCustomersDpt[years[p]]).length
        if (long > maxDpt) {
            maxDpt = long
        }
    }

    // premiere ligne dpt 
    let ligneUneDpt = []
    // Ajouter le nom de l'équipe
    for (let t = 0; t < years.length; t++) {
        ligneUneDpt.push(years[t], "Département", "Nombre client actif", "Nombre client potentiel", 'client actif / client potentiel %', "CA")
    }
    excelRowsDpt.push(ligneUneDpt)

    for (let k = 0; k < maxDpt; k++) {
        let rowDpt = []
        rowDpt.push('')
        for (let i = 0; i < years.length; i++) {
            let year = years[i]
            const sortedkey = Object.keys(allCustomersDpt[year]).sort()
                const key = sortedkey[k]
                let ratioDpt = ''
                if (potentiel[key] && potentiel[key] != 0) {
                    ratioDpt = (allCustomersDpt[year][key] / potentiel[key] * 100).toFixed(1) + '%'
                }
                // CA départements 
                let caD = 0 
                if (CADpts[year][key]) {
                    caD = CADpts[year][key]
                }
                rowDpt.push(key, allCustomersDpt[year][key], potentiel[key], ratioDpt,caD+"€", "")
            }
        excelRowsDpt.push(rowDpt)
    }


    
    dataBySheet.push(excelRowsDpt)


    /////////////////////////////////////////
    // Département client et potentiel regroupé par équipe (secteur)
    /////////////////////////////////////////
    let excelRowsDptequipe = []
    

    // premiere ligne dpt 
    ligneUneDpt = []
    // Ajouter le nom de l'équipe
    for (let t = 0; t < years.length; t++) {
        ligneUneDpt.push(years[t], "Secteur", "Département", "Nombre client actif", "Nombre client potentiel", 'client actif / client potentiel %', "CA")
    }
    excelRowsDptequipe.push(ligneUneDpt)

    const keyDpt = Object.keys(equipeParDepartement)

    for (let k = 0; k < keyDpt.length; k++) {
        let sumEquipe = {}
        let rowSum = []
        for (let t = 0; t < equipeParDepartement[keyDpt[k]].length; t++) {
            // somme 
            if (t == 0) {
                
            
            for (let h = 0; h < years.length; h++) {
                sumEquipe["Equipe"] = keyDpt[k]
                sumEquipe[years[h]] = {}
                sumEquipe[years[h]]['NbActif'] = 0
                sumEquipe[years[h]]['NbPotentiel'] = 0
                sumEquipe[years[h]]['CA'] = 0
            }
        }

            let rowDpt = []
            rowDpt.push('')
            for (let i = 0; i < years.length; i++) {
                
                let year = years[i]
                const key = equipeParDepartement[keyDpt[k]][t]
                
                let ratioDpt = ''
                if (potentiel[key] && potentiel[key] != 0) {
                    ratioDpt = (allCustomersDpt[year][key] / potentiel[key] * 100).toFixed(1)
                }
                // CA départements 
                let caD = 0
                if (CADpts[year][key]) {
                    caD = CADpts[year][key]
                }
                let nbActif = allCustomersDpt[year][key]
                let NbPotentiel = potentiel[key]
                sumEquipe[years[i]]['NbActif'] += nbActif
                sumEquipe[years[i]]['NbPotentiel'] += potentiel[key]
                
                sumEquipe[years[i]]['CA'] += caD
                rowDpt.push(keyDpt[k], key, nbActif, NbPotentiel, ratioDpt + "%", caD + "€", "")
            }
            excelRowsDptequipe.push(rowDpt)
        }
        for (let g = 0; g < years.length; g++) {
            let pourcent = (sumEquipe[years[g]]['NbActif']/sumEquipe[years[g]]['NbPotentiel']*100).toFixed(1)
            rowSum.push("", sumEquipe["Equipe"], "", sumEquipe[years[g]]['NbActif'], sumEquipe[years[g]]['NbPotentiel'],  pourcent+"%", sumEquipe[years[g]]['CA'] + "€")
        }
        excelRowsDptequipe.push(rowSum)
        excelRowsDptequipe.push([""])


    }



    dataBySheet.push(excelRowsDptequipe)



    
    //console.log(dataBySheet);
    return dataBySheet;
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

function generateOutput2(data){
    // Créer un nouveau classeur
    const newWorkbook = xlsx.utils.book_new()
    console.log("generate Excel...");

    let long = teams.length

    // pour toutes les équipes topAll
    for(let i = 0;i < long;i++){
    // Créer une nouvelle feuille
    const newWorksheet = xlsx.utils.aoa_to_sheet(data[i])
    // Ajouter la nouvelle feuille au classeur 
    xlsx.utils.book_append_sheet(newWorkbook,newWorksheet, teams[i].slice(0,30));
    }

    // National Stats
    // Créer une nouvelle feuille
    const worksheetNatio = xlsx.utils.aoa_to_sheet(data[long])
    
    // Ajouter la nouvelle feuille au classeur 
    xlsx.utils.book_append_sheet(newWorkbook,worksheetNatio, 'NATIONAL');

    // NewLost Stats
    // Créer une nouvelle feuille
    const worksheetNewLost = xlsx.utils.aoa_to_sheet(data[long+1])
    
    // Ajouter la nouvelle feuille au classeur 
    xlsx.utils.book_append_sheet(newWorkbook,worksheetNewLost, 'NewLost');

    // Département Stats
    // Créer une nouvelle feuille
    const worksheetDtp= xlsx.utils.aoa_to_sheet(data[long+2])
    
    // Ajouter la nouvelle feuille au classeur 
    xlsx.utils.book_append_sheet(newWorkbook,worksheetDtp, 'Dpt');

    // Département par équipe Stats
    // Créer une nouvelle feuille
    const worksheetDtpEqu= xlsx.utils.aoa_to_sheet(data[long+3])
    
    // Ajouter la nouvelle feuille au classeur 
    xlsx.utils.book_append_sheet(newWorkbook,worksheetDtpEqu, 'Dpt secteur');



    // Chemin du fichier de sortie
    const rd = Math.floor(Math.random() * 10000);
    const outputFilePath = "./excelGenerated/catalogueStats_"+ rd +".xlsx"
    console.log(outputFilePath);

    // Enregistrer le classeur dans un fichier Excel 
    xlsx.writeFile(newWorkbook,outputFilePath)

    return outputFilePath
}



async function colorEquipeVsNational(filePath) {
    const workbook = await XlsxPopulate.fromFileAsync(filePath);
    const alpha = ["D", "I", "N", "S", "X", "AC", "AH", "AM", "AR", "AW", "BA"]

    // pour chaque équipe
    for (let t = 0; t < teams.length; t++) {
        // la feuille de l'équipe en cours 
        const worksheet = workbook.sheet(t);
        // pour chaque année 
        for (let y = 0; y < years.length; y++) {
            // pour chaque produit
            for (let p = 0; p < topAll[0].length; p++) {
                let coordCell = alpha[y] + "" + (p + 2)
                const cell = worksheet.cell(coordCell)
                let cellValue = cell.value().toString().replace("%","")
                cellValue = parseFloat(cellValue)
                let natioValue = parseFloat(dataForColor[years[y]]["NATIONAL"][topAll[y][p]])

                if (cellValue < natioValue){
                    // rouge
                    cell.style({ fill: { type: 'solid', color: 'FC7E7E' } });
                }
                else if(cellValue == natioValue){
                    // Orange
                    cell.style({ fill: { type: 'solid', color: 'FFB185' } });
                }
                else{
                    // Vert
                    cell.style({ fill: { type: 'solid', color: '7EFC80' } });
                }
            }
        }
        
    }
    await workbook.toFileAsync(filePath);
}


/**
 * CONSTANTES
 */
let count = 0
const dataProut = getData()

const data = []
for (let i = 0; i < dataProut.length; i++) {
    if (typeof dataProut[i] != 'undefined') {
        data.push(dataProut[i])
    }
    
}
//console.log("Processing...");

const years = getYears().sort()
//console.log(years.length);

const teams = getTeams().sort()
//console.log(teams.length);

//const products = getProducts().sort()

const dpts = getDpts().sort()
//console.log(dpts.length);


//const top25 = getTop25PerYears()
const topAll = getTopPerYears()
const allCustomers = getAllCustomers()
const allCustomersProduct = getAllCustomersProduct()
//console.log(allCustomersProduct);
//const top25Price = getTop25PricePerYears()
const topAllPrice = getTopAllPricePerYears()

// DEPARTEMENT
const dataDPT = getDataEquipeDpt()
const potentiel = getPotentielDpts()
const allCustomersDpt = getAllCustomersDpt()
//const CADpts = getCADpts() // je pense pas nécessaire parce que je me suis trompé de fichier mais dans le doute...
const CADpts = getPricePerDtp()
console.log(CADpts);
const equipeParDepartement = getEquipeDpt()


// COMPARAISON POUR COLORER
const dataForColor = {}
for (let i = 0; i < years.length; i++) {
    dataForColor[years[i]] = {}
    for (let k = 0; k < teams.length; k++) {
        dataForColor[years[i]][teams[k]] = {}
    }
    dataForColor[years[i]]["NATIONAL"] = {}
}


// TEST
// pas normal que se soit à 0 mais normal parce que dans le fichier il y a 0 
//console.log(getCustomers('CD13 (75,92,93,94,95)','2021'))
//console.log(CADpts);

let sku = createDataDN()
// pour enlever tous les undefined qui me font un peu chier. UwU
for (let k = 0; k < sku.length; k++) {
    for (let q = 0; q < sku[k].length; q++) {
        for (let r = 0; r < sku[k][q].length; r++) {
            //console.log(sku[k][q][r]);
                if (typeof sku[k][q][r] === 'undefined') {
                    sku[k][q][r] = ''
                }
                if (typeof sku[k][q] === 'undefined') {
                    sku[k][q] = ''
                }
                if (typeof sku[k] === 'undefined') {
                    sku[k] = ''
                }
        }
    }
}

colorEquipeVsNational(generateOutput2(sku))

// fin du temps 
getTimeProcess(startTime, new Date())