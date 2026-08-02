// vybe-test: js/class_expression_named_anonymous/test_js_class_expression_super_access_inside_static_initializer_and_block
// origin: languages/js/tests/js/test_js_class_expression_named_anonymous.rs

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
    static baseValue = 10;
}

const Derived = class extends Base {
    static fromBase = super.baseValue;
    static fromBaseAndBlock;
    static {
        this.fromBaseAndBlock = super.baseValue + this.fromBase;
    }
};

__check(__line(Derived.fromBase), "10");
__check(__line(Derived.fromBaseAndBlock), "20");
