// vybe-test: js/operator_misc/in_operator_checks_own_and_inherited
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

const arr = [1, 2, 3];
__check(__line(0 in arr), "true");
__check(__line(10 in arr), "false");
__check(__line("length" in arr), "true");
const obj = { x: 1 };
__check(__line("x" in obj), "true");
__check(__line("toString" in obj), "true"); // inherited
