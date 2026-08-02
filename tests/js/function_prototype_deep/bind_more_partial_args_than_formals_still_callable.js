// vybe-test: js/function_prototype_deep/bind_more_partial_args_than_formals_still_callable
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

function take(a) { return a; } const fixed = take.bind(null, 1, 2, 3); __check(__line(fixed()), "1");
