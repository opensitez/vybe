// vybe-test: js/function_invocation_matrix/call_uses_explicit_receiver_for_plain_function
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

function label(prefix) {
    return prefix + ":" + this.name;
}
const obj = { name: "Ada" };
__check(__line(label.call(obj, "hi")), "hi:Ada");
