// vybe-test: js/json_deep/stringify_deeply_nested
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

const deep = { a: { b: { c: { d: 42 } } } };
const result = JSON.parse(JSON.stringify(deep));
__check(__line(result.a.b.c.d), "42");
