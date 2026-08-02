// vybe-test: js/operators_deep/logical_or_returns_first_truthy_or_last
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line(0 || false || 3), "3");  // first truthy
__check(__line(0 || false || ""), ""); // all falsy — last value
__check(__line(1 || 2 || 3), "1");      // first truthy
