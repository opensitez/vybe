// vybe-test: js/function_deep/call_sets_this_and_passes_args
// origin: languages/js/tests/js/test_function_deep.rs

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

function greet(greeting) { return greeting + " " + this.name; }
const obj = { name: "Bob" };
__check(__line(greet.call(obj, "Hi")), "Hi Bob");
