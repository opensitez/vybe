// vybe-test: js/delete_operator/delete_returns_true_on_success
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

const obj = { x: 1 };
__check(__line(delete obj.x), "true");
