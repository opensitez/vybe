// vybe-test: js/date_methods_deep/date_parse_returns_milliseconds
// origin: languages/js/tests/js/test_date_methods_deep.rs

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

const ms = Date.parse("2024-01-01T00:00:00.000Z");
console.log(typeof ms);
console.log(ms > 0);
// Verify it's correct
const d = new Date(ms);
console.log(d.getUTCFullYear());
