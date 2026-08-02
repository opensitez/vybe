// vybe-test: js/scope_closure_patterns/closure_counter_factory
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

function makeCounter(init = 0) {
    let count = init;
    return {
        increment: () => ++count,
        decrement: () => --count,
        reset: () => { count = init; },
        get: () => count,
    };
}
const c = makeCounter(10);
c.increment();
c.increment();
c.decrement();
__check(__line(c.get()), "11");
c.reset();
__check(__line(c.get()), "10");
