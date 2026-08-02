// vybe-test: js/structured_clone_patterns/structured_clone_primitive_wrappers_not_supported
// origin: languages/js/tests/js/test_structured_clone_patterns.rs

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

// Boolean, Number, String objects are cloneable
const n = new Number(42);
const clone = structuredClone(n);
__check(__line(+clone), "42");
