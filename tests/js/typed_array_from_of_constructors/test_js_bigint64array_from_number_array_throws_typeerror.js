// vybe-test: js/typed_array_from_of_constructors/test_js_bigint64array_from_number_array_throws_typeerror
// origin: languages/js/tests/js/test_js_typed_array_from_of_constructors.rs

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

try {
    BigInt64Array.from([1, 2]);
} catch (e) {
    __check(__line(e.name), "TypeError");
}
