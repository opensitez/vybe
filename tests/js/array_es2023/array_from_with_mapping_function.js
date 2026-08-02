// vybe-test: js/array_es2023/array_from_with_mapping_function
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

const arr = Array.from([1, 2, 3, 4], x => x * x);
__check(__line(arr.join(",")), "1,4,9,16");
