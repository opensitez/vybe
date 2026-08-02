// vybe-test: js/date_arithmetic_edge/date_set_full_year_mutates
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

const d=new Date(2020,0,1); d.setFullYear(2025); __check(__line(d.getFullYear()), "2025");
