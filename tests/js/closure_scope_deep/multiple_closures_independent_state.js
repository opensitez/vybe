// vybe-test: js/closure_scope_deep/multiple_closures_independent_state
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

function makeAdder(n) {
    return x => x + n;
}
const add5 = makeAdder(5);
const add10 = makeAdder(10);
__check(__line(add5(3)), "8");
__check(__line(add10(3)), "13");
__check(__line(add5(10) === add10(5)), "true");
