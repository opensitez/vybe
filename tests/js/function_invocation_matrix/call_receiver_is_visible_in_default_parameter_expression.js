// vybe-test: js/function_invocation_matrix/call_receiver_is_visible_in_default_parameter_expression
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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

function describe(prefix = this.tag) {
    __check(__line(prefix + ":" + this.tag), "ctx:ctx");
}
describe.call({ tag: "ctx" });
