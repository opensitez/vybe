// vybe-test: js/type_coercion_deep/to_primitive_all_converters_non_primitive_throws_typeerror
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
    valueOf() { return {}; },
    toString() { return {}; }
};
__check(__line(obj + 1), "[object Object]1");
