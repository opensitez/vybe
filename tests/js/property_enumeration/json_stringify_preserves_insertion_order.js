// vybe-test: js/property_enumeration/json_stringify_preserves_insertion_order
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const obj = { c: 3, a: 1, b: 2 };
const json = JSON.stringify(obj);
// JSON preserves insertion order
__check(__line(json), "{\"c\":3,\"a\":1,\"b\":2}");
