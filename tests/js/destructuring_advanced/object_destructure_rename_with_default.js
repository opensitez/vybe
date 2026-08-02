// vybe-test: js/destructuring_advanced/object_destructure_rename_with_default
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

const { x: px = 0, y: py = 0 } = { x: 5 };
__check(__line(px, py), "5 0");
