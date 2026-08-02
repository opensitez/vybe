// vybe-test: js/array_iteration_methods/copywithin_with_end
// origin: languages/js/tests/js/test_array_iteration_methods.rs

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

const arr = [1, 2, 3, 4, 5];
arr.copyWithin(1, 3, 4); // copy arr[3] to arr[1]
__check(__line(arr.join(",")), "1,4,3,4,5");
