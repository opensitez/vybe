// vybe-test: js/class_static_initialization_blocks/test_js_class_static_block_nested_class_definition
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

class Outer {
    static Inner;
    static {
        this.Inner = class {
            static name = "InnerClass";
        };
    }
}
__check(__line(Outer.Inner.name), "InnerClass");
