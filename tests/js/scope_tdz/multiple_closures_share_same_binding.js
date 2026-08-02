// vybe-test: js/scope_tdz/multiple_closures_share_same_binding
// origin: languages/js/tests/js/test_scope_tdz.rs

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

function makeCounter() {
    let n = 0;
    return {
        inc() { n++; },
        get() { return n; }
    };
}
const c = makeCounter();
c.inc(); c.inc(); c.inc();
__check(__line(c.get()), "3");
