// vybe-test: js/interop/test_a12_function_call_with_this_arg
// origin: languages/js/tests/js/js_interop_test.rs

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

function greet(greeting) {
            return greeting + " " + this.name;
        }
        let obj = { name: "Alice" };
        __check(__line(greet.call(obj, "Hello")), "Hello Alice");
