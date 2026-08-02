// vybe-test: js/json_patterns_deep/json_stringify_depth
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

const deep = { a: { b: { c: { d: 42 } } } };
const json = JSON.stringify(deep);
const parsed = JSON.parse(json);
__check(__line(parsed.a.b.c.d), "42");
