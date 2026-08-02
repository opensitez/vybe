// vybe-test: js/object_integrity_level_errors/delete_on_sealed_property_fails
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

const o=Object.seal({x:1}); __check(__line(delete o.x), "false");__check(__line(o.x), "1");
