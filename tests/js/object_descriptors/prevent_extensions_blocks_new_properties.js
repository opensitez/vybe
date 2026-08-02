// vybe-test: js/object_descriptors/prevent_extensions_blocks_new_properties
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const obj = { existing: 1 };
Object.preventExtensions(obj);
obj.newProp = 2;  // silently ignored
__check(__line(obj.existing), "1");
__check(__line("newProp" in obj), "false");
