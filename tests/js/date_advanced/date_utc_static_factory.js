// vybe-test: js/date_advanced/date_utc_static_factory
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

const ms = Date.UTC(2024, 0, 1);
const d = new Date(ms);
__check(__line(d.getUTCFullYear()), "2024");
__check(__line(d.getUTCMonth()), "0");
