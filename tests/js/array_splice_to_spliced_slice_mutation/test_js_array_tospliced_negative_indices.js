// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_tospliced_negative_indices
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

const arr = [1, 2, 3, 4];
const result = arr.toSpliced(-2, 1, 99);
__check(__line(result.join(",")), "1,2,99,4");
