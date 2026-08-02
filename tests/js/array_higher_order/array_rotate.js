// vybe-test: js/array_higher_order/array_rotate
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

function rotate(arr, n) {
    const k = ((n % arr.length) + arr.length) % arr.length;
    return [...arr.slice(k), ...arr.slice(0, k)];
}
__check(__line(rotate([1, 2, 3, 4, 5], 2).join(",")), "3,4,5,1,2");
__check(__line(rotate([1, 2, 3, 4, 5], 1).join(",")), "2,3,4,5,1");
