
const express = require('express');
const bodyParser = require('body-parser');
const { connect } = require('http2');
const { createConnection } = require('net');
const cors = require('cors');
require('dotenv').config();

// http://rbamvuates01:1017/
// const PATH = process.env.PATH ||'http://rbamvuates01:'
const PORT = process.env.PORT || 1017;

const app = express();

app.use(bodyParser.json());
app.use(cors());
app.listen(PORT, () => console.log(`Server running on port ${PORT}`)); // quitarlo en produccion???




//////////////////////// ROUTES ////////////////////////////////////////////////////////////////////////////////////////////////
app.get('/api/users', (req, res) => {
    const id = req.query.id;

    console.log(id);
    res.send(getAllApplications());
});

// 1. Api que devuelva conexiones
app.get('/api/Schemas', (req, res) => {
    res.send(getAllConections());
});
// POSTMAN http://localhost:3050/api/Schemas

// 2. Api que devuelva aplicaciones
app.get('/api/applications', (req, res) => {
    res.send(getAllApplications());
});
// POSTMAN http://localhost:3050/api/applications

// 3. Api que te devuelva las tablas y campos de una aplicacion /Nombre_Conexion/Nombre_Aplicacion
app.get('/api/querysApplication', (req, res) => {
    const application = req.query.application;

    res.send(filterByApp(application));
});
// POSTMAN http://localhost:3050/api/querysApplication?application=Listin

app.get('/api/Campos', (req, res) => {
    const schema = req.query.schema;
    const table = req.query.table;

    res.send(filterByTable(schema, table));
});
// POSTMAN http://localhost:3050/api/Campos?schema=HARVEST -  WE_IBERIA_LE_MART&table=CO_CALENDAR

app.get('/api/querys', (req, res) => {
    const schema = req.query.schema;
    const table = req.query.table;
    const field = req.query.field;

    if (field == "0") {
        res.send(filterByFieldTodo(schema, table));
    } else {
        res.send(filterByField(schema, table, field));
    }
});
// POSTMAN: http://localhost:3050/api/querys?schema=HARVEST -  WE_IBERIA_LE_MART&table=CO_CALENDAR&field=CALENDAR_LEVEL
// POSTMAN: http://localhost:3050/api/querys?schema=HARVEST -  WE_IBERIA_LE_MART&table=CO_CALENDAR&field=0

app.get('/api/Tables', (req, res) => {
    const schema = req.query.schema;

    res.send(filterByConection(schema));
});
// POSTMAN: http://localhost:3050/api/Tables?schema=HARVEST -  WE_IBERIA_LE_MART





//////////////////////// METHODS ///////////////////////////////////////////////////////////////////////////////////////////////

var XLSX = require("xlsx");
var alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
var result = ""

// const mainPath = "C:\\Users\\martm262\\Desktop\\VUE\\BACK-STRUCTURE\\conexiones\\";
const mainPath = "./connections/"; // Solo funciona cuando esta ejecutando el servidor (nodemon)


// excel.sheets devuelve ya todos los sheets no hace falta guardar sheet_names e ir recorriednolo
function getNumberRows(sheet) {

    var campos = []

    for (let i = 2; i < (Object.keys(sheet).length); i++) {
        try {
            var element = sheet["A" + i]["v"];
            campos.push(element);
        }
        catch {
            break;
        }
    }
    return campos;
};

function getNumberColumns(sheet) {

    var proyectos = [];

    for (let i = 2; i < (Object.keys(sheet).length); i++) {
        try {
            letra = printToLetter(i);
            element = sheet[letra + "1"]["v"];
            proyectos.push(element);
        }
        catch {
            break;
        }
    }
    return proyectos;
};

function printToLetter(number) {
    var charIndex = number % alphabet.length
    var quotient = number / alphabet.length
    if (charIndex - 1 == -1) {
        charIndex = alphabet.length
        quotient--;
    }
    result = alphabet.charAt(charIndex - 1) + result;
    if (quotient >= 1) {
        printToLetter(parseInt(quotient));
    } else {
        salida = result;
        result = ""
        return salida;
    }
};

function filterByApp(proyecto) {  // Es mas eficiente si en vez de recorrer todos para cada app se recorre para las dos a la vez comparando columnas con ambas
    conexiones = getAllConectionsList();
    salida1 = [];
    lista_scehmas = [];
    conexiones.forEach(conexion => {
        const excel = XLSX.readFile(mainPath + conexion + ".xlsx");
        var listaResultado = [];
        var sheets = excel.SheetNames;

        sheets.forEach(sheet => {

            current_sheet = excel.Sheets[sheet];

            valores = [];
            var numberOfColumns = getNumberColumns(current_sheet)
            for (let i = 2; i < (numberOfColumns.length + 2); i++) {
                letra = printToLetter(i);
                try {
                    element = current_sheet[letra + "1"]["h"]
                } catch {
                }
                if (element == proyecto) {
                    // buscar en la columna letra+todas las opciones
                    var numberOfRows = getNumberRows(current_sheet);
                    for (let j = 2; j < (numberOfRows.length + 2); j++) {
                        try {
                            campo = current_sheet["A" + j]["h"];
                            querys = current_sheet[letra + j]["h"];
                            valores.push({ "field": campo, "value": querys });
                        }
                        catch {
                        }
                    };
                    break;
                }
            }
            if (valores.length != 0) {//ver a que igualar para que no meta "table_options": []
                var resActual = { "table": sheet, "values": valores };
                listaResultado.push(resActual);
                
            }
        });
        if(listaResultado.length>0){
            var level1 = ({ "schema": conexion, "tables": listaResultado })
            lista_scehmas.push(level1);
        }
        
    });
    salida1 = ({ "project": proyecto, "schemas": lista_scehmas });
    jsonResultado = JSON.stringify(salida1);
    return jsonResultado;
};

function filterByConection(conexion) {
    const excel = XLSX.readFile(mainPath + conexion + ".xlsx");
    var nombreHojas = excel.SheetNames;
    jsonResultado = JSON.stringify(nombreHojas);
    return jsonResultado;
};

function filterByTable(conexion, tabla) {
    const excel = XLSX.readFile(mainPath + conexion + ".xlsx");
    var listaCampos = [];
    sheet = excel.Sheets[tabla];
    var numberOfColumns = getNumberColumns(sheet);
    var nomberOfRows = getNumberRows(sheet);

    for (let i = 2; i < (nomberOfRows.length + 2); i++) {
        try {
            is_used = false
            for (let j = 2; j < numberOfColumns.length + 2; j++) {
                try {
                    var letra = printToLetter(j);
                    // si hay algun elemento mas en la columna:
                    excel.Sheets[tabla][letra + i]["v"];
                    is_used = true;
                    break;
                }
                catch {
                }
            }
            if (is_used == true) {
                element = excel.Sheets[tabla]["A" + i]["v"];
                listaCampos.push(element);
            }
        }
        catch {
            break;
        }
    }
    jsonResultado = JSON.stringify(listaCampos);
    return jsonResultado;
};

function filterByTableEvenNotUsed(conexion, tabla) {
    const excel = XLSX.readFile(mainPath + conexion + ".xlsx");
    var listaCampos = [];
    for (let i = 2; i < (Object.keys(excel.Sheets[tabla]).length); i++) {
        try {
            element = excel.Sheets[tabla]["A" + i]["v"];
        }
        catch {
            break;
        }
        listaCampos.push(element);
    }
    jsonResultado = JSON.stringify(listaCampos);
    return jsonResultado;
};

function getAllConections() {
    salida = []
    var fs = require('fs');
    var files = fs.readdirSync(mainPath);

    files.forEach(element => {
        salida.push(element.replace(".xlsx", ""));
    });
    jsonResultado = JSON.stringify(salida);
    return jsonResultado;
};

function getAllConectionsList() {
    salida = []
    var fs = require('fs');
    var files = fs.readdirSync(mainPath);

    files.forEach(element => {
        salida.push(element.replace(".xlsx", ""));
    });
    return salida;
};

function getAllApplications() {
    conexiones = getAllConectionsList();
    var proyectos = [];

    conexiones.forEach(conexion => {
        const excel = XLSX.readFile(mainPath + conexion + ".xlsx");
        var sheets = excel.SheetNames; //array con todos los sheet names

        sheets.forEach(sheet => {
            current_sheet = excel.Sheets[sheet];
            var nomberOfColumns = getNumberColumns(current_sheet);

            for (let i = 2; i < (nomberOfColumns.length + 2); i++) {
                try {
                    element = excel.Sheets[sheet][printToLetter(i) + "1"]["v"];
                    if (!proyectos.includes(element)) {
                        proyectos.push(element);
                    }
                }
                catch {
                    break;
                }
            }
        });
    });
    proyectos.sort()
    jsonResultado = JSON.stringify(proyectos);
    return jsonResultado;
};

function filterByField(conexion, tabla, campo) {  // Es mas eficiente si en vez de recorrer todos para cada app se recorre para las dos a la vez comparando columnas con ambas

    const excel = XLSX.readFile(mainPath + conexion + ".xlsx");
    sheet = excel.Sheets[tabla];

    var listaResultado = [];

    // guardamos el indice del campo
    var indice = 0;
    var nomberOfRows = getNumberRows(sheet);

    for (let i = 2; i < (nomberOfRows.length + 2); i++) {
        try {
            element = excel.Sheets[tabla]["A" + i]["v"];
            if (element == campo) {
                indice = i;
                break;
            }
        }
        catch {
            break;
        }
    };

    var numberOfColumns = getNumberColumns(sheet);
    // recorremos la fila apara ese indice y guardamos la query y el campo si hay coincidencias
    for (let j = 2; j < numberOfColumns.length + 2; j++) {
        try {
            var letra = printToLetter(j);
            var element = sheet[letra + indice]["v"];
            var aplicacion = numberOfColumns[j - 2];

            listaResultado.push({ "field": campo, "project": aplicacion, "value": element });

        }  ////////CAMBIAR A LA UTLIMA FOTOOOOOOOOOOOOOOO 7/////////////////////
        catch {
        }
    };

    level1 = ({ "name": tabla, "values": listaResultado })
    level2 = ({ "schema": conexion, "tables": level1 })

    // });
    jsonResultado = JSON.stringify(level2);
    return jsonResultado;
};

function filterByFieldTodo(conexion, tabla) {  // Es mas eficiente si en vez de recorrer todos para cada app se recorre para las dos a la vez comparando columnas con ambas

    const excel = XLSX.readFile(mainPath + conexion + ".xlsx");
    sheet = excel.Sheets[tabla];

    var listaResultado = [];

    // guardamos el indice del campo
    var nomberOfRows = getNumberRows(sheet);

    for (let i = 2; i < (nomberOfRows.length + 2); i++) {
        try {
            current_field = excel.Sheets[tabla]["A" + i]["v"];

            var numberOfColumns = getNumberColumns(sheet);
            // recorremos la fila apara ese indice y guardamos la query y el campo si hay coincidencias
            for (let j = 2; j < numberOfColumns.length + 2; j++) {
                try {
                    var letra = printToLetter(j);
                    var current_query = sheet[letra + i]["v"];
                    var aplicacion = numberOfColumns[j - 2];

                    listaResultado.push({ "field": current_field, "project": aplicacion, "value": current_query });

                }  ////////CAMBIAR A LA UTLIMA FOTOOOOOOOOOOOOOOO 7/////////////////////
                catch {
                }
            };
        }
        catch {
            break;
        }
    };
    listaResultado.sort(function (a, b) {
        var nameA = a.project.toUpperCase(); // ignore upper and lowercase
        var nameB = b.project.toUpperCase(); // ignore upper and lowercase
        if (nameA < nameB) {
            return -1;
        }
        if (nameA > nameB) {
            return 1;
        }
        return 0;
    });
    level1 = ({ "name": tabla, "values": listaResultado })
    level2 = ({ "schema": conexion, "tables": level1 })

    // });
    jsonResultado = JSON.stringify(level2);
    return jsonResultado;
};





//////////// PROBAR: DESCARGA DE ARCHIVOS //////////////////////////////////////

// https://www.youtube.com/watch?v=VQOrkZrZXB8
// Y
// https://www.youtube.com/watch?v=pGtLTZHKCNo
