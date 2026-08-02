// vybe-test: js/closure_scope_deep_patterns/closure_over_mutable_reference
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

function makeAccumulator(initial = 0) {
    let total = initial;
    return {
        add(n) { total += n; return this; },
        subtract(n) { total -= n; return this; },
        result() { return total; },
    };
}
const acc = makeAccumulator(100);
acc.add(50).add(25).subtract(30);
__check(__line(acc.result()), "145");
