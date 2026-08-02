// vybe-test: js/operators_deep/logical_assignment_short_circuits_rhs
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

let value = 5;
let calls = 0;
value ||= (() => {
    calls++;
    return 9;
})();
let ready = 0;
ready ||= (() => {
    calls++;
    return 11;
})();
__check(__line(value), "5");
__check(__line(ready), "11");
__check(__line(calls), "1");
