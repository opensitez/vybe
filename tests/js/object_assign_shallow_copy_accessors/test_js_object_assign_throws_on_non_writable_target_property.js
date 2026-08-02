// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_throws_on_non_writable_target_property
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

const target = {};
Object.defineProperty(target, "fixed", { value: 10, writable: false });
try {
    Object.assign(target, { fixed: 20 });
} catch (e) {
    __check(__line("Assign ReadOnly Property TypeError"), "Assign ReadOnly Property TypeError");
}
