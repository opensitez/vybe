// vybe-test: js/json_patterns_deep/json_parse_reviver_all_keys
// origin: languages/js/tests/js/test_json_patterns_deep.rs

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

const calls = [];
const json = '{"a":1,"b":{"c":2}}';
JSON.parse(json, (key, val) => { calls.push(key); return val; });
// Reviver is called bottom-up: leaf keys first
__check(__line(calls.includes("a")), "true");
__check(__line(calls.includes("c")), "true");
__check(__line(calls[calls.length - 1]), ""); // root "" is last
