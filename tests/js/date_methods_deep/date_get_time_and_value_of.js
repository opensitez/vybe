// vybe-test: js/date_methods_deep/date_get_time_and_value_of
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

const d = new Date(12345678);
__check(__line(d.getTime()), "12345678");
__check(__line(d.valueOf()), "12345678");
__check(__line(+d), "12345678");
