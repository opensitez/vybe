// vybe-test: js/function_invocation_matrix/bind_preserves_first_bound_receiver
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

function show() {
    return this.name;
}
const bound = show.bind({ name: "first" });
__check(__line(bound.call({ name: "second" })), "first");
