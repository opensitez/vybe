// vybe-test: js/destructuring_comprehensive/array_destructuring_with_rest
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

const [first, second, ...rest] = [1, 2, 3, 4, 5];
__check(__line(first), "1");
__check(__line(second), "2");
__check(__line(rest.join(",")), "3,4,5");
