// vybe-test: js/ecma_variables/destructure_in_function_params
// origin: languages/js/tests/js/test_ecma_variables.rs

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

function greet({ name, age }) {
    __check(__line(name + " is " + age), "Bob is 25");
}
greet({ name: "Bob", age: 25 });
