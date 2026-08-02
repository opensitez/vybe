// vybe-test: js/coercion_toprimitive/object_to_boolean_always_truthy
// origin: languages/js/tests/js/test_coercion_toprimitive.rs

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

// All objects (even empty) are truthy
__check(__line(!!{}), "true");
__check(__line(!![]), "true");
__check(__line(!!new Boolean(false)), "true");
__check(__line(!!new Number(0)), "true");
