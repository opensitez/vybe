// vybe-test: js/destructuring_array_deep/destructure_with_rest
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

const [head, ...tail] = [1, 2, 3, 4, 5];
__check(__line(head), "1");
__check(__line(tail.join(",")), "2,3,4,5");
