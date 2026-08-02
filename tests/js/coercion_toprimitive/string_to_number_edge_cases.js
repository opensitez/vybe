// vybe-test: js/coercion_toprimitive/string_to_number_edge_cases
// origin: languages/js/tests/js/test_coercion_toprimitive.rs

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

__check(__line(Number(" 42 ")), "42"); // trims
__check(__line(Number("0x10")), "16"); // hex
__check(__line(Number("0o10")), "8"); // octal
__check(__line(Number("0b10")), "2"); // binary
__check(__line(Number("Infinity")), "Infinity");
__check(__line(Number("")), "0");
