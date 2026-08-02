// vybe-test: js/destructuring_patterns/array_destructure_rest_in_middle
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

// Rest must be last
const [first, ...remaining] = [1, 2, 3, 4, 5];
__check(__line(first), "1");
__check(__line(remaining.join(",")), "2,3,4,5");
