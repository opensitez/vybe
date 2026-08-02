// vybe-test: js/destructuring_advanced/mixed_destructure_object_with_arrays
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

const { nums: [a, b], tag } = { nums: [1, 2], tag: "ok" };
__check(__line(a, b, tag), "1 2 ok");
