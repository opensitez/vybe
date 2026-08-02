// vybe-test: js/json_serialization/json_roundtrip_types
// origin: languages/js/tests/js/test_json_serialization.rs

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

const original = { num: 42, str: "hello", bool: true, arr: [1,2,3], nil: null };
const json = JSON.stringify(original);
const parsed = JSON.parse(json);
__check(__line(parsed.num), "42");
__check(__line(parsed.str), "hello");
__check(__line(parsed.bool), "true");
__check(__line(parsed.arr.join(",")), "1,2,3");
__check(__line(parsed.nil), "null");
