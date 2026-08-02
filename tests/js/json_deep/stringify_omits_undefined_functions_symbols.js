// vybe-test: js/json_deep/stringify_omits_undefined_functions_symbols
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

const obj = {
    a: 1,
    b: undefined,
    c: () => {},
    d: Symbol("x"),
    e: "keep"
};
const result = JSON.parse(JSON.stringify(obj));
__check(__line(result.a), "1");
__check(__line("b" in result), "false");
__check(__line("c" in result), "false");
__check(__line("d" in result), "false");
__check(__line(result.e), "keep");
