// vybe-test: js/class_fields_advanced/field_initializer_order_within_class
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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
class Ordered {
    a = log.push("a") && 1;
    b = log.push("b") && 2;
    c = log.push("c") && 3;
    constructor() { log.push("ctor"); }
}
new Ordered();
__check(__line(log.join(",")), "a,b,c,ctor");
