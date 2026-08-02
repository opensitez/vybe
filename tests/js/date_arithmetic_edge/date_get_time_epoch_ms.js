// vybe-test: js/date_arithmetic_edge/date_get_time_epoch_ms
// origin: languages/js/tests/js/test_date_arithmetic_edge.rs

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

__check(__line(new Date("1970-01-01T00:00:00.000Z").getTime()), "0");
