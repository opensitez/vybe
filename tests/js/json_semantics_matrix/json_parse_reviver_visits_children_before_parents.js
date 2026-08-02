// vybe-test: js/json_semantics_matrix/json_parse_reviver_visits_children_before_parents
// origin: languages/js/tests/js/test_json_semantics_matrix.rs

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

const seen = [];
JSON.parse('{"outer":{"inner":1},"arr":[2]}', (key, value) => {
    seen.push(key === "" ? "<root>" : key);
    return value;
});
__check(__line(seen.join(",")), "inner,outer,0,arr,<root>");
