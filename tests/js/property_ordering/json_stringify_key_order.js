// vybe-test: js/property_ordering/json_stringify_key_order
// origin: languages/js/tests/js/test_property_ordering.rs

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

// JSON.stringify follows own enumerable insertion order (non-integer)
const obj = { b: 2, a: 1, c: 3 };
const json = JSON.stringify(obj);
__check(__line(json), "{\"b\":2,\"a\":1,\"c\":3}");
