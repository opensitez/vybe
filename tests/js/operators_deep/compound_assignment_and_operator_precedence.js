// vybe-test: js/operators_deep/compound_assignment_and_operator_precedence
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

let x = 20;
x += 10; // 30
x -= 5;  // 25
x *= 2;  // 50
x /= 5;  // 10
x %= 7;  // 3
x **= 2; // 9
__check(__line(x), "9");
