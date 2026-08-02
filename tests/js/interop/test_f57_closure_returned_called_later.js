// vybe-test: js/interop/test_f57_closure_returned_called_later
// origin: languages/js/tests/js/js_interop_test.rs

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

function makeMultiplier(factor) {
            return (x) => x * factor;
        }
        let triple = makeMultiplier(3);
        let quadruple = makeMultiplier(4);
        __check(__line(triple(10), quadruple(10)), "30 40");
