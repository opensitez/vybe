// vybe-test: js/function_call_apply_bind/call_to_borrow_method
// origin: languages/js/tests/js/test_function_call_apply_bind.rs

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

// Borrow Array.prototype.slice for array-like objects
function args() { return arguments; }
const argObj = args(1, 2, 3);
const arr = Array.prototype.slice.call(argObj);
console.log(Array.isArray(arr));
console.log(arr.join(","));
