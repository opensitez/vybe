// vybe-test: js/comprehensive/test_object_assign
// origin: languages/js/tests/js/js_comprehensive_test.rs

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

let a = { x: 1 };
        let b = { y: 2 };
        let c = Object.assign(a, b);
        __check(__line(c.x, c.y), "1 2");
