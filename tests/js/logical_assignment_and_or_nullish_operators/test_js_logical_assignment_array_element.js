// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_array_element
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

const arr = [0, null, 10];
arr[0] ||= 1;
arr[1] ??= 2;
arr[2] &&= 30;
__check(__line(arr.join(",")), "1,2,30");
