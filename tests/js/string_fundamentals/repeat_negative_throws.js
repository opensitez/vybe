// vybe-test: js/string_fundamentals/repeat_negative_throws
// origin: languages/js/tests/js/test_string_fundamentals.rs

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

console.log("ha".repeat(3));
try {
    console.log("x".repeat(-1));
} catch (e) {
    console.log(e.name);
}
