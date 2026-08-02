// vybe-test: js/date_utc_matrix/date_settime_updates_timestamp
// origin: languages/js/tests/js/test_date_utc_matrix.rs

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
d.setTime(1000);
__check(__line(d.getTime()), "1000");
__check(__line(d.getUTCSeconds()), "1");
