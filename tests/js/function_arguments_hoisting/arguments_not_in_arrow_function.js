// vybe-test: js/function_arguments_hoisting/arguments_not_in_arrow_function
// origin: languages/js/tests/js/test_function_arguments_hoisting.rs

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

function outer() {
    const inner = () => arguments[0]; // arrow captures outer's arguments
    return inner();
}
console.log(outer(42));
