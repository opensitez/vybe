// vybe-test: js/class_static_initialization_blocks/test_js_class_static_block_symbol_property_initialization
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

const symKey = Symbol("registry");
class Registry {
    static {
        this[symKey] = "RegisteredSymbolValue";
    }
    static getVal() { return this[symKey]; }
}
__check(__line(Registry.getVal()), "RegisteredSymbolValue");
