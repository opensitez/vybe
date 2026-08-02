// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_splice_omit_delete_count_deletes_to_end
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

const arr = [10, 20, 30, 40];
const removed = arr.splice(2);
__check(__line(arr.join(",") + "|removed=" + removed.join(",")), "10,20|removed=30,40");
