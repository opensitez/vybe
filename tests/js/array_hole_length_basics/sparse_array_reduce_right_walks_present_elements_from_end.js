// vybe-test: js/array_hole_length_basics/sparse_array_reduce_right_walks_present_elements_from_end
// origin: languages/js/tests/js/test_array_hole_length_basics.rs

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

const arr = [];
arr[1] = "b";
arr[3] = "d";
arr[5] = "f";
__check(__line(arr.reduceRight((acc, value) => acc + value, "")), "fdb");
