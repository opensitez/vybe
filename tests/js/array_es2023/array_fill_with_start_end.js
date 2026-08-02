// vybe-test: js/array_es2023/array_fill_with_start_end
// origin: languages/js/tests/js/test_array_es2023.rs

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

const arr = [0, 0, 0, 0, 0];
arr.fill(7, 1, 4);
__check(__line(arr.join(",")), "0,7,7,7,0");
