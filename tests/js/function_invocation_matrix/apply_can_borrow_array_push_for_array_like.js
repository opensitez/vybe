// vybe-test: js/function_invocation_matrix/apply_can_borrow_array_push_for_array_like
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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

const obj = { length: 0 };
Array.prototype.push.apply(obj, ["a", "b"]);
__check(__line(obj.length), "2");
__check(__line(obj[0] + obj[1]), "ab");
