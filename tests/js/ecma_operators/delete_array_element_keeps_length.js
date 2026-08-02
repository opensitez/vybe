// vybe-test: js/ecma_operators/delete_array_element_keeps_length
// origin: languages/js/tests/js/test_ecma_operators.rs

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

const arr = [10, 20, 30];
delete arr[1];
__check(__line(arr.length), "3");
__check(__line(1 in arr), "false");
__check(__line(arr[1]), "undefined");
