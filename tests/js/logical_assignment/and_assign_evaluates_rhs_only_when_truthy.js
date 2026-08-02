// vybe-test: js/logical_assignment/and_assign_evaluates_rhs_only_when_truthy
// origin: languages/js/tests/js/test_logical_assignment.rs

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

let calls = 0;
let x = true;
x &&= (calls++, false);
__check(__line(calls), "1");
let y = false;
y &&= (calls++, true);
__check(__line(calls), "1");
