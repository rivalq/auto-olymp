var axios = require("axios");
var jsdom = require("jsdom");
var https = require("https");
var http = require("http");
var fs = require("fs");
var path = require("path");
var root = __dirname;
var PDFParser = require("pdf2json");
const cheerio = require("cheerio");

const httpAgent = new http.Agent({ keepAlive: true });

module.exports = function () {
    const { JSDOM } = jsdom;
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
        let temp = arr[i]["dob"];
        let vals = temp.split("/");
        if (vals.length > 0 && vals[0].length == 1) vals[0] = "0" + vals[0];
        if (vals.length > 1 && vals[1].length == 1) vals[1] = "0" + vals[1];
        if (vals.length > 2) arr[i]["dob"] = vals[0] + "/" + vals[1] + "/" + vals[2];
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

    const remove_empty = (arr) => {
        let temp = [];
        for (let i = 0; i < arr.length; i++) {
            if (arr[i] != "") temp.push(arr[i]);
        }
        return temp;
    };

    const extract = (obj, page) => {
        const dom = new JSDOM(page);
        let title = dom.window.document.title;
        let process_array = [];
        if (title[0] == "I") {
            obj["Valid"] = "NO";
            write_to_sheet(obj);
        } else {
            const $ = cheerio.load(page);
            let first_name = $("#counts > div > div > div > div > div > div > div > div > div > div > div.row.text-left > div:nth-child(2) > div:nth-child(1) > div:nth-child(1) > label.font-weight-bold.text-danger").text();
            let middle_name = $("#counts > div > div > div > div > div > div > div > div > div > div > div.row.text-left > div:nth-child(2) > div:nth-child(1) > div:nth-child(2) > label.font-weight-bold.text-danger").text();
            let last_name = $("#counts > div > div > div > div > div > div > div > div > div > div > div.row.text-left > div:nth-child(2) > div:nth-child(1) > div:nth-child(3) > label.font-weight-bold.text-danger").text();
            let mobile_no = $("#counts > div > div > div > div > div > div > div > div > div > div > div.row.text-left > div:nth-child(2) > div:nth-child(2) > div:nth-child(1) > label.font-weight-bold.text-danger").text();
            let email = $("#counts > div > div > div > div > div > div > div > div > div > div > div.row.text-left > div:nth-child(2) > div:nth-child(2) > div:nth-child(2) > label.font-weight-bold.text-danger").text();
            let subject = $("#counts > div > div > div > div > div > div > div > div > div > div > div.row.text-left > div:nth-child(3) > div > div:nth-child(3) > label.font-weight-bold.text-danger.text-center").text();
            let language = $("#counts > div > div > div > div > div > div > div > div > div > div > div.row.text-left > div:nth-child(3) > div > div:nth-child(4) > label.font-weight-bold.text-danger").text();
            let fee_paid = $("#counts > div > div > div > div > div > div > div > div > div > div > div:nth-child(11) > div:nth-child(1) > label.font-weight-bold.text-success").text();
            obj["Valid"] = "YES";
            obj["First Name"] = first_name;
            obj["Middle Name"] = middle_name;
            obj["Last Name"] = last_name;
            obj["Mobile No."] = mobile_no;
            obj["Email"] = email;
            obj["Subject"] = subject;
            obj["Language"] = language;
            obj["Fee Paid"] = fee_paid;
            write_to_sheet(obj);
        }
        return process_array;
    };
    const make_request = async (data, obj) => {
        var config = {
            method: "post",
            url: "https://stureg.ioqmexam.in/Login",
            headers: {
                Cookie: "",
                ...data.getHeaders(),
            },
            data: data,
        };
        const result = await axios.get("https://stureg.ioqmexam.in/Login");
        const $ = cheerio.load(result.data);
        const token = $('input[name="__RequestVerificationToken"]').val();
        data.append("__RequestVerificationToken", token);
        config["headers"]["Cookie"] = result.headers["set-cookie"];
        await axios(config)
            .then(function (response) {
                extract(obj, response.data);
            })
            .catch(function (error) {
                console.log("Error", error);
                console.log("Error occured for current entry");
                obj["valid"] = "NO";
                write_to_sheet(obj);
            });
    };

    const func = async () => {
        for (let i = 0; i < 1; i++) {
            var FormData = require("form-data");
            var data = new FormData();
            var obj = new Object();
            obj = {
                "Registration No.": "",
                DOB: "",
                Valid: "",
            };
            obj["Registration No."] = arr[i]["reg_no"];
            obj["DOB"] = arr[i]["dob"];
            data.append("RegNo", arr[i]["reg_no"]);
            data.append("DOB", arr[i]["dob"]);
            console.log("Processing Entry ", i + 1);
            const result = await make_request(data, obj);
            //await delay(1000);
        }
    };

    func().then((data) => {
        commit();
        console.log("Completed !");
    });
};
