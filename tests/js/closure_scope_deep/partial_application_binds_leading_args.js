// vybe-test: js/closure_scope_deep/partial_application_binds_leading_args
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

function partial(fn, ...preset) {
    return (...rest) => fn(...preset, ...rest);
}

const multiply = (a, b) => a * b;
const triple = partial(multiply, 3);
__check(__line(triple(5)), "15");
__check(__line(triple(10)), "30");
