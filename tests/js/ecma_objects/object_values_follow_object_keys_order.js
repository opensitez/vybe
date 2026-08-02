// vybe-test: js/ecma_objects/object_values_follow_object_keys_order
// origin: languages/js/tests/js/test_ecma_objects.rs

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

const obj = { a: 10, b: 20, c: 30 };
__check(__line(Object.values(obj).join(",")), "10,20,30");
