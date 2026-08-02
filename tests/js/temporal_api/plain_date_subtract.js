// vybe-test: js/temporal_api/plain_date_subtract
// origin: languages/js/tests/js/test_temporal_api.rs

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

const d = new Date(2024, 2, 10);
d.setDate(d.getDate() - 5);
__check(__line(d.getDate()), "5");
