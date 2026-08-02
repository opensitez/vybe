// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_hints_in_various_operations
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
    [Symbol.toPrimitive](hint) {
        log.push(hint);
        return hint === "string" ? "str" : 10;
    }
};

String(obj); // string
Number(obj); // number
obj + 5;     // default
obj == 10;   // default
__check(__line(log.join(",")), "string,number,default,default");
