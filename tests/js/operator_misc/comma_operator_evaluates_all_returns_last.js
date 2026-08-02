// vybe-test: js/operator_misc/comma_operator_evaluates_all_returns_last
// origin: languages/js/tests/js/test_operator_misc.rs

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

let a = 0, b = 0;
const result = (a++, b++, a + b);
__check(__line(result), "2");
__check(__line(a), "1");
__check(__line(b), "1");
