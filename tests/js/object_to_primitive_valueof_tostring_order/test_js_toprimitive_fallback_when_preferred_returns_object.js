// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_fallback_when_preferred_returns_object
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

const log = [];
const obj = {
    valueOf() { log.push("valueOf"); return {}; }, // Returns object, not primitive!
    toString() { log.push("toString"); return "fallbackStr"; }
};
const res = Number(obj);
__check(__line(res + "|" + log.join(",")), "NaN|valueOf,toString");
