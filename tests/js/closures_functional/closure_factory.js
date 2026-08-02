// vybe-test: js/closures_functional/closure_factory
// origin: languages/js/tests/js/test_closures_functional.rs

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

function makeCounter(start) {
    let count = start;
    return {
        next() { return count++; },
        reset() { count = start; }
    };
}
let c = makeCounter(10);
__check(__line(c.next()), "10");
__check(__line(c.next()), "11");
__check(__line(c.next()), "12");
c.reset();
__check(__line(c.next()), "10");
