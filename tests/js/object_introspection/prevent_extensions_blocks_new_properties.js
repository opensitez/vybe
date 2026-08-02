// vybe-test: js/object_introspection/prevent_extensions_blocks_new_properties
// origin: languages/js/tests/js/test_object_introspection.rs

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

const obj = { existing: "yes" };
Object.preventExtensions(obj);
obj.newProp = "no";
__check(__line(obj.existing), "yes");
__check(__line(obj.newProp), "undefined");
