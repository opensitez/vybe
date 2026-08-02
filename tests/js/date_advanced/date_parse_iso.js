// vybe-test: js/date_advanced/date_parse_iso
// origin: languages/js/tests/js/test_date_advanced.rs

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

const ms = Date.parse("2024-06-15T00:00:00.000Z");
__check(__line(typeof ms), "number");
__check(__line(ms > 0), "true");
