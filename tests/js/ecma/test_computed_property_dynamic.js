// vybe-test: js/ecma/test_computed_property_dynamic
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

let obj = {};
        for (let i = 0; i < 3; i++) {
            obj["key" + i] = i * 10;
        }
        console.log(obj.key0, obj.key1, obj.key2);
