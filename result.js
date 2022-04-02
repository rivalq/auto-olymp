var FormData = require("form-data");
var axios = require("axios");
var jsdom = require("jsdom");
const { JSDOM } = jsdom;
var fs = require("fs");
var path = require("path");

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
const write_to_sheet = (obj) => {
    output.push(obj);
};
const scrape_data = (obj, dom) => {
    obj["valid"] = "YES";
    let header = dom.window.document.getElementsByClassName("col-sm-4");
    let child = header[3].childNodes;
    obj["Name"] = child[3].textContent + child[5].textContent;
    child = header[5].childNodes;
    obj["Roll Number"] = child[3].textContent;
    let middle = dom.window.document.getElementsByClassName("col-sm-3");
    child = middle[0].childNodes;
    obj["Total Attended"] = child[3].textContent;
    child = middle[1].childNodes;
    obj["Total Not Attended"] = child[3].textContent;
    child = middle[2].childNodes;
    obj["Total Correct"] = child[3].textContent;
    child = middle[3].childNodes;
    obj["Total Wrong"] = child[3].textContent;
    middle = dom.window.document.getElementsByClassName("col-sm-12");
    obj["Tentative Score"] = middle[3].childNodes[3].textContent;
    let table = dom.window.document.getElementsByTagName("tr");
    for (let i = 1; i < table.length; i++) {
        let child = table[i].childNodes;
        obj[i] = stripspacs(child[7].textContent);
    }
    write_to_sheet(obj);
};

const get_result = async (obj, data) => {
    var config = {
        method: "post",
        url: "https://reg.ioqexam.in/Login",
        headers: {
            Cookie: ".ASPXAUTH=1C5072A12029D5BDEDBA4A8B2BF34BCF93857C12508D6E08AE261C8E09F777DCCDE6BEF372BE71DE3F2030333E8FFCFF27B64BF33553EB86C03D0916B49E1AAF66BB4E0F1FCBB789BFDADB1BCDA8BB5E486C1A7951B5FD2DFA06F2EA3676FCA5177C0105AA679CC7B7A487FE7220EF7F; ASP.NET_SessionId=vw1mld10s5lgaabp4nbts1mr",
            ...data.getHeaders(),
        },
        data: data,
    };
    await axios(config)
        .then(function (response) {
            const dom = new JSDOM(response.data);
            let title = dom.window.document.title;
            if (title[0] == "I") {
                obj["valid"] = "NO";
                write_to_sheet(obj);
            }
        })
        .catch(function (error) {
            obj["valid"] = "NO";
            write_to_sheet(obj);
        });
    if (obj["valid"] != "NO") {
        let url = `https://reg.ioqexam.in/Application/Results?Regno=${obj["RegNo"]}&ioqSubject=IOQM`;
        var config2 = {
            method: "get",
            url: url,
            headers: {
                Cookie: ".ASPXAUTH=1C5072A12029D5BDEDBA4A8B2BF34BCF93857C12508D6E08AE261C8E09F777DCCDE6BEF372BE71DE3F2030333E8FFCFF27B64BF33553EB86C03D0916B49E1AAF66BB4E0F1FCBB789BFDADB1BCDA8BB5E486C1A7951B5FD2DFA06F2EA3676FCA5177C0105AA679CC7B7A487FE7220EF7F; ASP.NET_SessionId=vw1mld10s5lgaabp4nbts1mr",
            },
        };
        await axios(config2)
            .then(function (response) {
                const dom = new JSDOM(response.data);
                let title = dom.window.document.title;
                if (title[0] == "D") {
                    obj["valid"] = "NO";
                    write_to_sheet(obj);
                } else {
                    scrape_data(obj, dom);
                }
            })
            .catch(function (error) {
                obj["valid"] = "NO";
                write_to_sheet(obj);
            });
    }
};

const main = async () => {
    for (let i = 0; i < arr.length; i++) {
        let regno = arr[i]["Reg No."];
        let dob = arr[i]["DOB"];
        var data = new FormData();
        data.append("RegNo", regno);
        data.append("DOB", dob);
        var obj = {
            RegNo: "",
            DOB: "",
            Name: "",
            valid: "",
            "Roll Number": "",
            "Total Attended": "",
            "Total Not Attended": "",
            "Total Correct": "",
            "Total Wrong": "",
            "Tentative Score": "",
            1: "",
            2: "",
            3: "",
            4: "",
            5: "",
            6: "",
            7: "",
            8: "",
            9: "",
            10: "",
            11: "",
            12: "",
        };
        obj["RegNo"] = regno;
        obj["DOB"] = dob;
        console.log("Processing Entry ", i + 1);
        await get_result(obj, data);
    }
};

main().then(() => {
    commit();
    console.log("Completed !");
});
