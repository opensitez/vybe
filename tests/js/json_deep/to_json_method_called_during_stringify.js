// vybe-test: js/json_deep/to_json_method_called_during_stringify
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
    value: 42,
    secret: "hidden",
    toJSON() { return { value: this.value }; }
};
const result = JSON.parse(JSON.stringify(obj));
__check(__line(result.value), "42");
__check(__line("secret" in result), "false");
