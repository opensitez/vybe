// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_null_or_undefined_target_throws_typeerror
// origin: languages/js/tests/js/test_js_object_assign_shallow_copy_accessors.rs

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

try {
    Object.assign(null, { a: 1 });
} catch (e) {
    __check(__line("Assign Null Target TypeError"), "Assign Null Target TypeError");
}
