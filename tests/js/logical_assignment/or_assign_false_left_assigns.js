// vybe-test: js/logical_assignment/or_assign_false_left_assigns
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

let x = 0;
x ||= 99;
__check(__line(x), "99");
