// vybe-test: js/date/date_from_timestamp
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

let d = new Date(0);
__check(__line(d.getUTCFullYear()), "1970");
__check(__line(d.getUTCMonth()), "0");
__check(__line(d.getUTCDate()), "1");
