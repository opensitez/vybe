// vybe-test: js/class_extends_super_constructor_call/test_js_class_super_spread_arguments_forwarding
// origin: languages/js/tests/js/test_js_class_extends_super_constructor_call.rs

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

class Base {
    constructor(...args) {
        this.sum = args.reduce((a, b) => a + b, 0);
    }
}
class Sub extends Base {
    constructor(multiplier, ...nums) {
        super(...nums);
        this.total = this.sum * multiplier;
    }
}
__check(__line(new Sub(10, 1, 2, 3).total), "60");
