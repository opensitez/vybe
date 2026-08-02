// vybe-test: js/json_semantics_matrix/json_parse_then_stringify_sorts_integer_like_keys
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

const json = JSON.stringify(JSON.parse('{"10":1,"2":2,"a":3}'));
__check(__line(json), "{\"2\":2,\"10\":1,\"a\":3}");
