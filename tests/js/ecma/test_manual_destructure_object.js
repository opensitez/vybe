// vybe-test: js/ecma/test_manual_destructure_object
// origin: languages/js/tests/js/js_ecma_test.rs

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

let obj = { x: 10, y: 20 };
        let x = obj.x;
        let y = obj.y;
        __check(__line(x + y), "30");
