// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_date_object_default_hint_is_string
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
const d = new Date(0);
d[Symbol.toPrimitive] = function(hint) {
    log.push(hint);
    return Date.prototype[Symbol.toPrimitive].call(this, hint);
};
const res = d + 10;
__check(__line(log.join(",") + "|" + (typeof res)), "default|string");
