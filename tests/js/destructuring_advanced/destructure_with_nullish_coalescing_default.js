// vybe-test: js/destructuring_advanced/destructure_with_nullish_coalescing_default
// origin: languages/js/tests/js/test_destructuring_advanced.rs

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

const config = {};
const { timeout = 5000, retries = 3 } = config;
__check(__line(timeout, retries), "5000 3");
