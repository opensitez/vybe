// vybe-test: js/property_descriptor_deep/prevent_extensions_blocks_new_props
// origin: languages/js/tests/js/test_property_descriptor_deep.rs

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

const obj = { a: 1 };
Object.preventExtensions(obj);
obj.b = 2; // silently fails
__check(__line(obj.b), "undefined");
__check(__line(Object.isExtensible(obj)), "false");
