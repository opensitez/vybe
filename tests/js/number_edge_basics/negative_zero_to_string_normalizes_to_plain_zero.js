// vybe-test: js/number_edge_basics/negative_zero_to_string_normalizes_to_plain_zero
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

__check(__line((-0).toString()), "0");
