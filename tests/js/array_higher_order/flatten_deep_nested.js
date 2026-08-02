// vybe-test: js/array_higher_order/flatten_deep_nested
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

const nested = [1, [2, [3, [4, [5]]]]];
__check(__line(nested.flat(Infinity).join(",")), "1,2,3,4,5");
__check(__line(nested.flat(2).join(",")), "1,2,3,4,5");
