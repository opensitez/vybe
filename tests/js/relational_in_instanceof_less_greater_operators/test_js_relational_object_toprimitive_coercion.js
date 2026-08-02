// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_relational_object_toprimitive_coercion
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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

const obj1 = { [Symbol.toPrimitive]: () => 10 };
const obj2 = { valueOf: () => 20 };
__check(__line(obj1 < obj2), "true");
