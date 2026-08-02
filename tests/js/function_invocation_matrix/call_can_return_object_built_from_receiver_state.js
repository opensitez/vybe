// vybe-test: js/function_invocation_matrix/call_can_return_object_built_from_receiver_state
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

function make(key) {
    return { seen: this[key] };
}
const ctx = { value: 7 };
__check(__line(make.call(ctx, "value").seen), "7");
