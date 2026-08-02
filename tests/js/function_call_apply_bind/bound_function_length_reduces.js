// vybe-test: js/function_call_apply_bind/bound_function_length_reduces
// origin: languages/js/tests/js/test_function_call_apply_bind.rs

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

function f(a, b, c) { return a + b + c; }
__check(__line(f.length), "3");        // 3
const g = f.bind(null, 1);    // partial: 1 arg bound
__check(__line(g.length), "2");        // 2
const h = f.bind(null, 1, 2); // 2 args bound
__check(__line(h.length), "1");        // 1
