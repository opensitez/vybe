// vybe-test: js/function_arguments_hoisting/arguments_index_access
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

function sum() {
    let s = 0;
    for (let i = 0; i < arguments.length; i++) s += arguments[i];
    return s;
}
console.log(sum(1, 2, 3, 4));
