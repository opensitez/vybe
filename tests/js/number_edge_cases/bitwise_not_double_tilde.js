// vybe-test: js/number_edge_cases/bitwise_not_double_tilde
// origin: languages/js/tests/js/test_number_edge_cases.rs

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

// ~~x converts to int32 (truncates)
__check(__line(~~3.9), "3");
__check(__line(~~-3.9), "-3");
__check(__line(~~"42"), "42");
__check(__line(~~null), "0");
