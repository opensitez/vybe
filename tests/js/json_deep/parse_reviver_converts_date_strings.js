// vybe-test: js/json_deep/parse_reviver_converts_date_strings
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

const json = '{"created":"2024-01-15","value":42}';
const result = JSON.parse(json, (key, value) => {
    if (key === "created") return new Date(value).getFullYear();
    return value;
});
__check(__line(result.created), "2024");
__check(__line(result.value), "42");
