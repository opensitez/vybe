// vybe-test: js/symbol_to_primitive_coercion_hint/test_js_symbol_to_primitive_array_join_bypasses_to_primitive
// origin: languages/js/tests/js/test_js_symbol_to_primitive_coercion_hint.rs

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
    [Symbol.toPrimitive]() { return "Bypassed"; },
    toString() { return "CalledToString"; }
};
__check(__line([obj].join("")), "CalledToString");
