// vybe-test: js/delete_operator/delete_var_does_not_work
// origin: languages/js/tests/js/test_delete_operator.rs

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

var x = 1;
const result = delete x;
__check(__line(result), "false");    // false — vars are non-configurable globals
__check(__line(typeof x), "number");  // still exists
