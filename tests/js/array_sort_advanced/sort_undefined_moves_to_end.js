// vybe-test: js/array_sort_advanced/sort_undefined_moves_to_end
// origin: languages/js/tests/js/test_array_sort_advanced.rs

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

const arr = [3, undefined, 1, undefined, 2];
arr.sort();
// undefined values go to the end
__check(__line(arr[arr.length - 1] === undefined), "true");
__check(__line(arr[arr.length - 2] === undefined), "true");
