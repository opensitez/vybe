// vybe-test: js/features/test_closure
// origin: languages/js/tests/js/js_features_test.rs

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
            return () => { count = count + 1; return count; };
        }
        let c = makeCounter();
        __check(__line(c(), c(), c()), "1 2 3");
