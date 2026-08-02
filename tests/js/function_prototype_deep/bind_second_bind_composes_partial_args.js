// vybe-test: js/function_prototype_deep/bind_second_bind_composes_partial_args
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

function concat(a, b, c) { return "" + a + b + c; } const step = concat.bind(null, "a"); const done = step.bind(null, "b"); __check(__line(done("c")), "abc");
