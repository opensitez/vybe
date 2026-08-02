// vybe-test: js/symbol_to_primitive_coercion_hint/test_js_symbol_to_primitive_not_callable_throws_typeerror
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

const obj = { [Symbol.toPrimitive]: "not_a_function" };
try {
    +obj;
} catch (e) {
    __check(__line("ToPrimitive Not Callable TypeError"), "ToPrimitive Not Callable TypeError");
}
