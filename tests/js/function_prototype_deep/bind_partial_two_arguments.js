// vybe-test: js/function_prototype_deep/bind_partial_two_arguments
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

function mul(a, b, c) { return a * b * c; } const timesSix = mul.bind(null, 2, 3); __check(__line(timesSix(4)), "24");
