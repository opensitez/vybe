// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_splice_sparse_array_holes
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

const sparse = [1, , 3, 4];
const removed = sparse.splice(1, 2);
__check(__line(sparse.join(",") + "|removedLen=" + removed.length + "|removedHole=" + !(0 in removed)), "1,4|removedLen=2|removedHole=true");
