// vybe-test: js/destructuring_patterns/array_destructure_default_from_previous_element
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

const [a = 10, b = a * 2] = [5];
__check(__line(a), "5");
__check(__line(b), "10"); // b = a * 2 = 10
