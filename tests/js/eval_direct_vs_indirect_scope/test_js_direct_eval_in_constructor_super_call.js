// vybe-test: js/eval_direct_vs_indirect_scope/test_js_direct_eval_in_constructor_super_call
// origin: languages/js/tests/js/test_js_eval_direct_vs_indirect_scope.rs

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

class Base { constructor(x) { this.x = x; } }
class Derived extends Base {
    constructor(x) {
        eval("super(x * 2)");
    }
}
__check(__line(new Derived(10).x), "20");
