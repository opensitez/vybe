// vybe-test: js/date_utc_matrix/date_local_setmilliseconds_overflow_rolls_second
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

const d = new Date(2024, 0, 1, 1, 2, 3, 999);
d.setMilliseconds(1000);
__check(__line(d.getSeconds()), "4");
__check(__line(d.getMilliseconds()), "0");
