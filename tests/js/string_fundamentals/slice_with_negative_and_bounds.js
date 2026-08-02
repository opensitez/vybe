// vybe-test: js/string_fundamentals/slice_with_negative_and_bounds
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

const s = "abcdef";
__check(__line(s.slice(-2)), "ef");
__check(__line(s.slice(2, 4)), "cd");
__check(__line(s.slice(-10, 2)), "ab");
