// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_bigint_returned_by_symbol_to_primitive
// origin: languages/js/tests/js/test_js_object_to_primitive_valueof_tostring_order.rs

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
    [Symbol.toPrimitive]() { return 100n; }
};
__check(__line((obj == 100n) + "|" + (typeof obj)), "true|object");
