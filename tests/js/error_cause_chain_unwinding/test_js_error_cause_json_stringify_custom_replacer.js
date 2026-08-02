// vybe-test: js/error_cause_chain_unwinding/test_js_error_cause_json_stringify_custom_replacer
// origin: languages/js/tests/js/test_js_error_cause_chain_unwinding.rs

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

const cause = new Error("InnerMsg");
const err = new Error("OuterMsg", { cause });
const json = JSON.stringify(err, (key, value) => {
    if (value instanceof Error) {
        return { name: value.name, message: value.message, cause: value.cause };
    }
    return value;
});
__check(__line(json.includes("OuterMsg") && json.includes("InnerMsg")), "true");
