// vybe-test: js/closure_scope_deep/counter_via_closure
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

function makeCounter(start = 0) {
    let count = start;
    return {
        increment() { return ++count; },
        decrement() { return --count; },
        reset() { count = start; return count; },
        value() { return count; }
    };
}
const c = makeCounter(10);
__check(__line(c.increment()), "11");
__check(__line(c.increment()), "12");
__check(__line(c.decrement()), "11");
__check(__line(c.reset()), "10");
