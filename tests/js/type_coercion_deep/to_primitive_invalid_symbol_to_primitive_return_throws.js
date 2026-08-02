// vybe-test: js/type_coercion_deep/to_primitive_invalid_symbol_to_primitive_return_throws
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

const bad = {
    [Symbol.toPrimitive]() {
        return {};
    }
};

try {
    console.log(bad == 1);
} catch (e) {
    console.log(e.name);
}

console.log(bad + "x");
