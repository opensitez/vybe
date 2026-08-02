// vybe-test: js/module_patterns/module_with_state
// origin: languages/js/tests/js/test_module_patterns.rs

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

const Counter = (() => {
    let count = 0;
    return {
        inc: () => ++count,
        dec: () => --count,
        reset: () => { count = 0; return count; },
        value: () => count,
    };
})();
Counter.inc();
Counter.inc();
Counter.inc();
Counter.dec();
__check(__line(Counter.value()), "2");
Counter.reset();
__check(__line(Counter.value()), "0");
