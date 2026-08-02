// vybe-test: js/missing_features/object_spread_basic
// origin: languages/js/tests/js/js_missing_features_test.rs

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

let a = { x: 1, y: 2 };
        let b = { ...a, z: 3 };
        __check(__line(b.x), "1");
        __check(__line(b.y), "2");
        __check(__line(b.z), "3");
