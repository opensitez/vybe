// vybe-test: js/function_invocation_matrix/apply_receiver_used_in_default_parameter
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

function greet(name = this.name) {
    return "hi " + name;
}
__check(__line(greet.apply({ name: "Ada" }, [])), "hi Ada");
