// vybe-test: js/ecma_strings/json_roundtrip
// origin: languages/js/tests/js/test_ecma_strings.rs

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

const orig = { a: 1, b: "hello", c: true };
const s = JSON.stringify(orig);
const parsed = JSON.parse(s);
__check(__line(parsed.a), "1");
__check(__line(parsed.b), "hello");
__check(__line(parsed.c), "true");
