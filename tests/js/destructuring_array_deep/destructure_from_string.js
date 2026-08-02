// vybe-test: js/destructuring_array_deep/destructure_from_string
// origin: languages/js/tests/js/test_destructuring_array_deep.rs

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

const [a, b, c] = "hello";
__check(__line(a), "h");
__check(__line(b), "e");
__check(__line(c), "l");
