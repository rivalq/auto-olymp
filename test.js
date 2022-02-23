const axios = require("axios");
const http = require("http");
const https = require("https");

const path = require("path");
const fs = require("fs");
const axios2 = axios.create({
    timeout: 1000,
    httpAgent: new http.Agent({ keepAlive: true, maxSockets: 1, keepAliveMsecs: 10000 }),
    httpsAgent: new https.Agent({ keepAlive: true, maxSockets: 1, keepAliveMsecs: 10000 }),
});

var FormData = require("form-data");
var data = new FormData();
let regn = "IOQ22012716Kar";
let date = "08/09/2004";
data.append("RegNo", regn);
data.append("DOB", date);

var config = {
    method: "post",
    url: "https://reg.ioqexam.in/Login",
    headers: {
        Cookie: ".ASPXAUTH=774029711337A91A3BE74FACF4740098FA669944108DBD4ADCDC49A745F6E555DCB0466FC5CDBEB959FCED47196476B45556C687C2AA8DF5CDA35CF01E2AB4DCE1DB21B07F20AB48BE24762250C87A17A0FEE60AA132B4DC84DCC3426CCADF020AB8BD7DD4DB1492290FCEC599F095BA; ASP.NET_SessionId=xq0rfvdn2cz0wye0qrgvj2bc",
        ...data.getHeaders(),
    },
    data: data,
};
let url = "https://reg.ioqexam.in/Application/GetHalltickets?Regno=" + regn + "&subid=" + "IOQM" + "&rType=HAL";

var config2 = {
    method: "get",
    url: "https://reg.ioqexam.in/Application/GetHalltickets?Regno=" + regn + "&subid=" + "IOQM" + "&rType=HAL",
    headers: {
        //Cookie: ".ASPXAUTH=774029711337A91A3BE74FACF4740098FA669944108DBD4ADCDC49A745F6E555DCB0466FC5CDBEB959FCED47196476B45556C687C2AA8DF5CDA35CF01E2AB4DCE1DB21B07F20AB48BE24762250C87A17A0FEE60AA132B4DC84DCC3426CCADF020AB8BD7DD4DB1492290FCEC599F095BA; ASP.NET_SessionId=xq0rfvdn2cz0wye0qrgvj2bc",
    },
    responseType: "stream",
};

axios(config)
    .then(function (response) {
        console.log(response);
        /*axios(config2).then((res) => {
            const pat = path.resolve("Admit_Cards", "IOQM" + "_" + regn + ".pdf");
            //console.log(pat);
            const writeStream = fs.createWriteStream(pat);
            res.data.pipe(writeStream);

            writeStream.on("finish", () => {
                writeStream.close();

                //const pdfParser = new PDFParser(this, 1);
                //pdfParser.on("pdfParser_dataReady", (pdfData) => {
                //    if (subject == "IOQM") {
                //        process2(obj, pdfParser.getRawTextContent(pdfData));
                //    }
                //    resolve(true);
                //});
            });
        });*/
    })
    .catch(function (error) {
        //console.log(error);
    });
