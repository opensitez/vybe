// vybe-test: js/closure_scope_deep_patterns/iife_private_scope
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

const counter = (() => {
    let count = 0;
    return {
        increment() { count++; },
        decrement() { count--; },
        value() { return count; }
    };
})();
counter.increment();
counter.increment();
counter.increment();
counter.decrement();
__check(__line(counter.value()), "2");
