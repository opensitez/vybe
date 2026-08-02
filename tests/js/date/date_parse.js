// vybe-test: js/date/date_parse
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

let ts = Date.parse("2024-01-01T00:00:00Z");
__check(__line(typeof ts), "number");
__check(__line(ts > 0), "true");
