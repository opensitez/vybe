// vybe-test: js/json_deep/stringify_replacer_array_allows_only_listed_keys
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

const obj = { a: 1, b: 2, c: 3, d: 4 };
const json = JSON.stringify(obj, ["a", "c"]);
const result = JSON.parse(json);
__check(__line(result.a), "1");
__check(__line("b" in result), "false");
__check(__line(result.c), "3");
