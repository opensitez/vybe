// vybe-test: js/change_array_by_copy_to_reversed_to_spliced_to_sorted_with/test_js_array_with_array_like_object
// origin: languages/js/tests/js/test_js_change_array_by_copy_to_reversed_to_spliced_to_sorted_with.rs

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

const arrayLike = { 0: 10, 1: 20, length: 2 };
const updated = Array.prototype.with.call(arrayLike, 0, 99);
__check(__line(Array.isArray(updated) + "|" + updated.join(",")), "true|99,20");
