// vybe-test: js/destructuring_array_deep/destructure_default_values
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

const [a = 10, b = 20, c = 30] = [1, 2];
__check(__line(a), "1");
__check(__line(b), "2");
__check(__line(c), "30");
