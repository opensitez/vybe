// vybe-test: js/functional_patterns_deep/once_fn_called_only_once
// origin: languages/js/tests/js/test_functional_patterns_deep.rs

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

function once(fn) {
    let called = false, result;
    return function(...args) {
        if (!called) { called = true; result = fn(...args); }
        return result;
    };
}
let n = 0;
const inc = once(() => ++n);
__check(__line(inc()), "1");
__check(__line(inc()), "1");
__check(__line(inc()), "1");
__check(__line(n), "1");
