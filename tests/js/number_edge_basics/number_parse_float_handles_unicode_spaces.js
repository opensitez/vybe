// vybe-test: js/number_edge_basics/number_parse_float_handles_unicode_spaces
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

__check(__line(parseFloat("\t  42.5\n")), "42.5");
__check(__line(parseFloat("-10.5foo")), "-10.5");
