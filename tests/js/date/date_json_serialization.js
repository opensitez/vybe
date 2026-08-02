// vybe-test: js/date/date_json_serialization
// origin: languages/js/tests/js/test_date.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

let d = new Date("2024-06-15T00:00:00Z");
let json = JSON.stringify({ date: d });
__check(__line(json.includes("2024")), "true");
