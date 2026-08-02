// vybe-test: js/json_patterns_deep/json_stringify_replacer_function
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

const obj = { a: 1, b: 2, c: 3, secret: "hidden" };
const result = JSON.stringify(obj, (key, val) => {
    if (key === "secret") return undefined;
    return val;
});
__check(__line(result.includes("secret")), "false");
__check(__line(result.includes("a")), "true");
