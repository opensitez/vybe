// vybe-test: js/destructuring_patterns/destructure_default_not_applied_to_null
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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

// Default only applied when value is undefined, not null
const { x = "default" } = { x: null };
__check(__line(x), "null");
