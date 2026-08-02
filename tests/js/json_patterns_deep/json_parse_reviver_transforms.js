// vybe-test: js/json_patterns_deep/json_parse_reviver_transforms
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

const json = '{"name":"Alice","birthDate":"2000-01-01"}';
const parsed = JSON.parse(json, (key, val) => {
    if (key === "birthDate") return new Date(val);
    return val;
});
__check(__line(parsed.name), "Alice");
__check(__line(parsed.birthDate instanceof Date), "true");
