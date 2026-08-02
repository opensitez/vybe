// vybe-test: js/date_advanced/date_from_iso_string
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

const d = new Date("2024-06-15T00:00:00.000Z");
__check(__line(d.getUTCFullYear()), "2024");
__check(__line(d.getUTCMonth()), "5");
__check(__line(d.getUTCDate()), "15");
