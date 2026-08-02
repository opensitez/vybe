// vybe-test: js/delete_operator/delete_property_descriptor_rejects_undeletable_property_in_strict_eval
// origin: languages/js/tests/js/test_delete_operator.rs

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
    eval('"use strict"; const o = Object.create(null); Object.defineProperty(o, "locked", { value: 1, configurable: false}); delete o.locked;');
    console.log("no-error");
} catch (e) {
    console.log(e.name);
}
