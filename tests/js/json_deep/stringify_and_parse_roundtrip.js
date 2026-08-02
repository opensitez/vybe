// vybe-test: js/json_deep/stringify_and_parse_roundtrip
// origin: languages/js/tests/js/test_json_deep.rs

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

const obj = { name: "Alice", age: 30, active: true };
const json = JSON.stringify(obj);
const parsed = JSON.parse(json);
__check(__line(parsed.name), "Alice");
__check(__line(parsed.age), "30");
__check(__line(parsed.active), "true");
