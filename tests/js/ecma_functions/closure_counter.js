// vybe-test: js/ecma_functions/closure_counter
// origin: languages/js/tests/js/test_ecma_functions.rs

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
    let count = 0;
    return function() {
        count += 1;
        return count;
    };
}
const counter = makeCounter();
__check(__line(counter()), "1");
__check(__line(counter()), "2");
__check(__line(counter()), "3");
