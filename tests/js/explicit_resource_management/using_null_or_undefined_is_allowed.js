// vybe-test: js/explicit_resource_management/using_null_or_undefined_is_allowed
// origin: languages/js/tests/js/test_explicit_resource_management.rs

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

let ok = true;
try {
    using r = null;
    using s = undefined;
} catch {
    ok = false;
}
__check(__line(ok), "true");
