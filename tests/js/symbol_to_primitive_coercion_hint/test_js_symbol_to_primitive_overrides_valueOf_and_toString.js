// vybe-test: js/symbol_to_primitive_coercion_hint/test_js_symbol_to_primitive_overrides_valueOf_and_toString
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
    valueOf() { return 10; },
    toString() { return "10"; },
    [Symbol.toPrimitive](hint) {
        return 999;
    }
};
__check(__line(+obj + "|" + String(obj)), "999|999");
