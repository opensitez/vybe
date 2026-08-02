// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_symbol_to_primitive_overrides_both
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
    valueOf() { log.push("valueOf"); return 10; },
    toString() { log.push("toString"); return "str"; },
    [Symbol.toPrimitive](hint) { log.push("toPrimitive:" + hint); return 99; }
};
const res = obj + 1;
__check(__line(res + "|" + log.join(",")), "100|toPrimitive:default");
