// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_slice_shallow_copy_subset
// origin: languages/js/tests/js/test_js_array_splice_to_spliced_slice_mutation.rs

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

const arr = ["a", "b", "c", "d", "e"];
const sliced = arr.slice(1, 4);
__check(__line(sliced.join(",") + "|orig=" + arr.join(",")), "b,c,d|orig=a,b,c,d,e");
