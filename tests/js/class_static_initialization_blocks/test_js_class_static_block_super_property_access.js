// vybe-test: js/class_static_initialization_blocks/test_js_class_static_block_super_property_access
// origin: languages/js/tests/js/test_js_class_static_initialization_blocks.rs

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
    static baseValue = "BaseVal";
}
class Sub extends Base {
    static subValue;
    static {
        this.subValue = super.baseValue + "_Extended";
    }
}
__check(__line(Sub.subValue), "BaseVal_Extended");
