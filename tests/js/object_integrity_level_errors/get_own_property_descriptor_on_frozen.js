// vybe-test: js/object_integrity_level_errors/get_own_property_descriptor_on_frozen
// origin: languages/js/tests/js/test_object_integrity_level_errors.rs

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

const d=Object.getOwnPropertyDescriptor(Object.freeze({x:1}),"x"); __check(__line(d.writable), "false");
