// vybe-test: js/fixes/test_has_own_property
// origin: languages/js/tests/js/js_fixes_test.rs

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

let obj = { x: 1, y: 2 };
        __check(__line(obj.hasOwnProperty("x")), "true");
