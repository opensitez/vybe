// vybe-test: js/json_deep/date_to_json_returns_iso_string
// origin: languages/js/tests/js/test_json_deep.rs

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

const d = new Date(0); // epoch
const json = JSON.stringify({ date: d });
__check(__line(json.includes("1970")), "true");
