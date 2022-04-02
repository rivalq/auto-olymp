var axios = require("axios");
var jsdom = require("jsdom");
var https = require("https");
var http = require("http");
var fs = require("fs");
var path = require("path");
var root = __dirname;
var PDFParser = require("pdf2json");
const httpAgent = new http.Agent({ keepAlive: true });

module.exports = function () {
    const { JSDOM } = jsdom;
    const delay = (ms) => new Promise((res) => setTimeout(res, ms));

    const reader = require("xlsx");
    const file = reader.readFile("data.xlsx", { cellDates: true });

    let arr = [];
    let output = [];
    let process_array = [];
    const sheets = file.SheetNames;

    for (let i = 0; i < 1; i++) {
        const temp = reader.utils.sheet_to_json(file.Sheets[file.SheetNames[i]], { raw: false });

        temp.forEach((res) => {
            arr.push(res);
        });
    }

    for (let i = 0; i < arr.length; i++) {
        let temp = arr[i]["DOB"];
        let vals = temp.split("/");
        if (vals.length > 0 && vals[0].length == 1) vals[0] = "0" + vals[0];
        if (vals.length > 1 && vals[1].length == 1) vals[1] = "0" + vals[1];
        if (vals.length > 2) arr[i]["DOB"] = vals[1] + "/" + vals[0] + "/" + vals[2];
    }

    const commit = () => {
        try {
            fs.unlinkSync(path.resolve("output.xlsx"));
        } catch (err) {}
        fs.writeFile(path.resolve("output.xlsx"), "", function () {
            const file = reader.readFile("output.xlsx");
            const ws = reader.utils.json_to_sheet(output);
            reader.utils.book_append_sheet(file, ws, "Output");
            reader.writeFile(file, "output.xlsx");
        });
    };

    const write_to_sheet = (data) => {
        output.push(data);
    };

    const stripspacs = (str) => {
        let dat = str.split("\n");
        let t = "";
        for (let i = 0; i < dat.length; i++) {
            let arr = dat[i].split(" ");
            for (let i = 0; i < arr.length; i++) {
                t += arr[i];
            }
        }

        return t;
    };

    const getFileUrl = (obj, subject) => {
        let regn = obj["Registration No."];
        let url = "https://reg.ioqexam.in/Application/GetHalltickets?Regno=" + regn + "&subid=" + subject + "&rType=HAL";
        return url;
    };

    const remove_empty = (arr) => {
        let temp = [];
        for (let i = 0; i < arr.length; i++) {
            if (arr[i] != "") temp.push(arr[i]);
        }
        return temp;
    };

    const process = (obj, card) => {
        card = card.split("\n");
        obj["ROLL No."] = stripspacs(card[5].split(":")[1]);
        let str = card[9].split(" ");
        str = remove_empty(str);
        obj["Gender"] = str[2];
        obj["Class"] = str[8];
        str = remove_empty(card[12].split(" "));
        str.pop();
        str.pop();
        for (let i = 4; i < str.length; i++) {
            obj["School"] += str[i] + " ";
        }

        for (let i = 16; i < str.length; i++) {
            let str = remove_empty(card[i].split(" "));
            if (str[0] == "Subject") break;
            for (let j = 0; j < str.length; j++) {
                if (str[j] == "Student" || str[j] == "across") break;
                obj["Center"] += str[j] + " ";
            }
        }
        write_to_sheet(obj);
    };

    const process2 = (obj, card) => {
        card = card.split("\n");
        pending = new Object();

        for (let i = 0; i < card.length; i++) {
            try {
                card[i] = card[i].split("\r")[0];
                let str = card[i].split(" ");
                if (str[0] == "School" && obj["School"] == "") {
                    str = card[i].split(":");
                    obj["School"] = str[1];
                }
                if (str[0] == "Gender" && obj["Gender"] == "") {
                    str = card[i].split(":");
                    obj["Gender"] = str[1];
                }
                if (str[0] == "Class" && obj["Class"] == "") {
                    str = card[i].split(":");
                    obj["Class"] = str[1];
                }
                if (str[0] == "OFFICE" && i + 1 < card.length && obj["ROLL No."] == "") {
                    str = card[i + 1].split("\r")[0];
                    obj["ROLL No."] = str;
                }
            } catch (err) {}
        }
        for (let i = 0; i < card.length; i++) {
            let str = card[i];
            if (str == "NAME AND LOCATION OF THE TEST CENTRE AND CONTACT DETAILS : " && obj["Center"] == "") {
                let j = i + 1;
                while (j < card.length) {
                    let temp = card[j];
                    if (temp == "Attach your recent") {
                        break;
                    }
                    obj["Center"] += temp;
                    j++;
                }
                break;
            }
        }

        write_to_sheet(obj);
    };

    const extract = (data, page) => {
        const dom = new JSDOM(page);
        let title = dom.window.document.title;
        let process_array = [];
        if (title[0] == "I") {
            data["valid"] = "NO";
            process_array.push([data, -1]);
        } else {
            data["valid"] = "YES";
            let names = dom.window.document.getElementsByClassName("col-sm-4");
            if (names.length > 3 && names[3].childNodes.length > 3) data["First Name"] = names[3].childNodes[3].textContent;
            if (names.length > 5 && names[5].childNodes.length > 3) data["Last Name"] = names[5].childNodes[3].textContent;
            let arr = dom.window.document.getElementsByClassName("col-sm-6");
            if (arr.length > 0 && arr[0].childNodes.length > 5) data["Mobile"] = arr[0].childNodes[5].textContent;
            if (arr.length > 1 && arr[1].childNodes.length > 3) data["Email"] = arr[1].childNodes[3].textContent;
            let subjects = ["PrevSubJ1", "PrevSubP1", "PrevSubB1", "PrevSubC1", "PrevSubA1", "PrevSubM1"];
            let pref = ["PrevCen1", "PrevCen2", "PrevCen3"];
            let temp = ["First Preference", "Second Preference", "Third Preference"];
            for (let i = 0; i < pref.length; i++) {
                let arr = dom.window.document.getElementById(pref[i]).childNodes;
                if (arr.length > 3) {
                    let center = stripspacs(arr[3].textContent);
                    data[temp[i]] = center;
                }
            }
            for (let i = 0; i < subjects.length; i++) {
                let arr = dom.window.document.getElementById(subjects[i]).childNodes;
                if (arr.length > 5) {
                    let temp = stripspacs(arr[5].textContent);
                    let subject = stripspacs(arr[1].textContent);
                    data[subject] = temp;
                    if (temp != "0") {
                        console.log(subject);
                        process_array.push([JSON.parse(JSON.stringify(data)), subject]);
                    }
                }
            }
            //write_to_sheet(data);
        }
        return process_array;
    };
    const make_request = async (data, obj) => {
        var config = {
            method: "post",
            url: "https://reg.ioqexam.in/Login",
            headers: {
                Cookie: ".ASPXAUTH=1C5072A12029D5BDEDBA4A8B2BF34BCF93857C12508D6E08AE261C8E09F777DCCDE6BEF372BE71DE3F2030333E8FFCFF27B64BF33553EB86C03D0916B49E1AAF66BB4E0F1FCBB789BFDADB1BCDA8BB5E486C1A7951B5FD2DFA06F2EA3676FCA5177C0105AA679CC7B7A487FE7220EF7F; ASP.NET_SessionId=vw1mld10s5lgaabp4nbts1mr",
                ...data.getHeaders(),
            },
            data: data,
        };

        const aux_func = async (vals) => {
            let obj = vals[0];
            let subject = vals[1];
            if (subject == -1) {
                write_to_sheet(obj);
                return;
            }
            let url = getFileUrl(obj, subject);
            let regn = obj["Registration No."];
            var dir = "./Admit_Cards";
            if (!fs.existsSync(dir)) {
                fs.mkdirSync(dir);
            }
            var config2 = {
                method: "get",
                url: url,
                headers: {
                    Cookie: ".ASPXAUTH=1C5072A12029D5BDEDBA4A8B2BF34BCF93857C12508D6E08AE261C8E09F777DCCDE6BEF372BE71DE3F2030333E8FFCFF27B64BF33553EB86C03D0916B49E1AAF66BB4E0F1FCBB789BFDADB1BCDA8BB5E486C1A7951B5FD2DFA06F2EA3676FCA5177C0105AA679CC7B7A487FE7220EF7F; ASP.NET_SessionId=vw1mld10s5lgaabp4nbts1mr",
                },
                responseType: "stream",
            };
            await axios(config2).then((res) => {
                const pat = path.resolve("Admit_Cards", subject + "_" + regn + ".pdf");
                const writeStream = fs.createWriteStream(pat);
                res.data.pipe(writeStream);
                writeStream.on("finish", () => {
                    writeStream.close();
                    const pdfParser = new PDFParser(this, 1);
                    pdfParser.on("pdfParser_dataReady", (pdfData) => {
                        if (subject == "IOQM") {
                            process2(obj, pdfParser.getRawTextContent(pdfData));
                        }
                    });
                    pdfParser.loadPDF(pat);
                });
            });
        };
        let arr = [];
        await axios(config)
            .then(function (response) {
                arr = extract(obj, response.data);
            })
            .catch(function (error) {
                console.log("error");
                obj["valid"] = "NO";
                write_to_sheet(obj);
            });
        for (let i = 0; i < arr.length; i++) {
            await aux_func(arr[i]);
        }
    };

    //make_request(data, obj);
    const func = async () => {
        for (let i = 0; i < arr.length; i++) {
            var FormData = require("form-data");
            var data = new FormData();
            var obj = new Object();
            obj = {
                "Registration No.": "",
                DOB: "",
                "ROLL No.": "",
                valid: 0,
                "First Name": "",
                "Last Name": "",
                Gender: "",
                School: "",
                Mobile: "",
                Email: "",
                Class: "",
                Center: "",
                IOQM: "",
                IOQA: "",
                IOQJ: "",
                IOQP: "",
                IOQC: "",
                IOQB: "",
                "First Preference": "",
                "Second Preference": "",
                "Third Preference": "",
            };
            obj["Registration No."] = arr[i]["Reg No."];
            obj["DOB"] = arr[i]["DOB"];
            data.append("RegNo", arr[i]["Reg No."]);
            data.append("DOB", arr[i]["DOB"]);
            console.log("Processing Entry ", i + 1);
            const result = await make_request(data, obj);
            //await delay(2000);
        }
    };

    func().then((data) => {
        commit();
        console.log("Completed !");
    });
};
