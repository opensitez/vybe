// vybe-test: js/operators_deep/in_operator_with_symbol_property_key
// origin: languages/js/tests/js/test_operators_deep.rs

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

const token = Symbol("token");
const obj = { [token]: 123 };
__check(__line(token in obj), "true");
__check(__line(Object.hasOwn(obj, token)), "true");
