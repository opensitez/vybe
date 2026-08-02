// vybe-test: js/object_integrity_level_errors/keys_on_sealed_returns_existing
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

__check(__line(Object.keys(Object.seal({x:1,y:2})).sort().join(",")), "x,y");
