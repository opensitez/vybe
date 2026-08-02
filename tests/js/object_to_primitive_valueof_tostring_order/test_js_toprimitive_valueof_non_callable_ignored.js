// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_valueof_non_callable_ignored
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
    valueOf: "not_a_function",
    toString: () => "validStr"
};
__check(__line(Number(obj)), "NaN"); // Non-callable valueOf is ignored, falls back to toString!
