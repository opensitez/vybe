// vybe-test: js/scope_closure_patterns/closure_over_mutable_variable
// origin: languages/js/tests/js/test_scope_closure_patterns.rs

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

function makeAdder(base) {
    return function(n) {
        base += n;
        return base;
    };
}
const add = makeAdder(10);
__check(__line(add(5)), "15");
__check(__line(add(3)), "18");
__check(__line(add(-2)), "16");
