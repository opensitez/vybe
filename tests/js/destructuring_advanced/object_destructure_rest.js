// vybe-test: js/destructuring_advanced/object_destructure_rest
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

const { x, ...rest } = { x: 1, y: 2, z: 3 };
__check(__line(x), "1");
__check(__line(Object.keys(rest).sort().join(",")), "y,z");
