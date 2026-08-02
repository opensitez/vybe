// vybe-test: js/interop/test_b19_object_assign_merges
// origin: languages/js/tests/js/js_interop_test.rs

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
        let c = { z: 3 };
        let merged = Object.assign({}, a, b, c);
        __check(__line(merged.x, merged.y, merged.z), "1 2 3");
