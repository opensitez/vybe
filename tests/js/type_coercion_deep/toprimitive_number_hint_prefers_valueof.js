// vybe-test: js/type_coercion_deep/toprimitive_number_hint_prefers_valueof
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

const obj = {
    valueOf() { return 42; },
    toString() { return "str"; }
};
__check(__line(obj - 0), "42");    // number hint → valueOf
__check(__line(`${obj}`), "str");  // string hint → toString
