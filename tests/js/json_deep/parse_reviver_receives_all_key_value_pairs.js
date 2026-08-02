// vybe-test: js/json_deep/parse_reviver_receives_all_key_value_pairs
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

const keys = [];
JSON.parse('{"a":1,"b":{"c":2}}', (key, value) => {
    if (key !== "") keys.push(key);
    return value;
});
__check(__line(keys.sort().join(",")), "a,b,c");
