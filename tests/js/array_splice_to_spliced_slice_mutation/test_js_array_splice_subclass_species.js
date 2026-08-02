// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_splice_subclass_species
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

class CustomArray extends Array {}
const ca = new CustomArray(1, 2, 3);
const removed = ca.splice(0, 1);
__check(__line(ca.join(",") + "|removedIsCustom=" + (removed instanceof CustomArray)), "2,3|removedIsCustom=true");
