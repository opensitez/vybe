// vybe-test: js/destructuring_advanced/array_destructure_from_function
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

function getCoords() { return [10, 20]; }
const [x, y] = getCoords();
__check(__line(x + y), "30");
