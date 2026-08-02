// vybe-test: js/date_advanced/date_to_iso_string_epoch
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

const d = new Date(0);
__check(__line(d.toISOString()), "1970-01-01T00:00:00.000Z");
