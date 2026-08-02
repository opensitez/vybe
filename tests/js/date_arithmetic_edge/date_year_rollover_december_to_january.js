// vybe-test: js/date_arithmetic_edge/date_year_rollover_december_to_january
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

const d=new Date(2023,11,31); d.setDate(d.getDate()+1); __check(__line(d.getMonth()), "0");
