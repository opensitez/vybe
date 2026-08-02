// vybe-test: js/type_coercion_deep/to_boolean_truthy_falsy
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

__check(__line(Boolean("")), "false");
__check(__line(Boolean("0")), "true");
__check(__line(Boolean(0)), "false");
__check(__line(Boolean(NaN)), "false");
__check(__line(Boolean({})), "true");
__check(__line(Boolean([])), "true");
__check(__line(Boolean(Symbol("coercion"))), "true");
