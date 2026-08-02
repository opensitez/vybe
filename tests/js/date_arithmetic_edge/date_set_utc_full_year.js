// vybe-test: js/date_arithmetic_edge/date_set_utc_full_year
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

const d=new Date(0); d.setUTCFullYear(2021); __check(__line(d.getUTCFullYear()), "2021");
