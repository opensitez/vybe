// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_tospliced_copies_holes_as_undefined
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

const sparse = [1, , 3];
const result = sparse.toSpliced(0, 0);
__check(__line(result.length + "|" + result.map(x => String(x)).join(",")), "3|1,undefined,3");
