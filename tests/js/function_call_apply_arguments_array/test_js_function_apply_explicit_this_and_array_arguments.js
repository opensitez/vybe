// vybe-test: js/function_call_apply_arguments_array/test_js_function_apply_explicit_this_and_array_arguments
// origin: languages/js/tests/js/test_js_function_call_apply_arguments_array.rs

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

function greet(greeting, punctuation) {
    return `${greeting} ${this.name}${punctuation}`;
}
const user = { name: "Bob" };
__check(__line(greet.apply(user, ["Hi", "?"])), "Hi Bob?");
