// vybe-test: js/destructuring_array_deep/destructure_rename_and_default
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

const { x: myX = 10, y: myY = 20 } = { x: 5 };
__check(__line(myX), "5");
__check(__line(myY), "20");
