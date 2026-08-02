// vybe-test: js/array_higher_order/array_from_with_mapping_function_and_this_arg
// origin: languages/js/tests/js/test_array_higher_order.rs

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

const arr = Array.from([1, 2, 3], v => v * 3);
__check(__line(arr.join(",")), "3,6,9");
