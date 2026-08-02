// vybe-test: js/eval_dynamic_code/new_function_can_use_closures_via_outer_function
// origin: languages/js/tests/js/test_eval_dynamic_code.rs

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

function makeAdder(n) {
    // Can't capture n via new Function, so pass it as arg
    return new Function("x", "return x + " + n);
}
const add5 = makeAdder(5);
console.log(add5(10));
