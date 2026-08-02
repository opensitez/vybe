// vybe-test: js/coercion_modern/array_isarray_edge_cases
// origin: languages/js/tests/js/test_coercion_modern.rs

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

__check(__line(Array.isArray([])), "true");
__check(__line(Array.isArray(new Array())), "true");
__check(__line(Array.isArray({})), "false");
__check(__line(Array.isArray("string")), "false");
__check(__line(Array.isArray(Array.of(1, 2))), "true");
