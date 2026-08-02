// vybe-test: js/number_edge_basics/parse_float_reads_exponent_notation
// origin: languages/js/tests/js/test_number_edge_basics.rs

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

__check(__line(parseFloat("1.25e2")), "125");
