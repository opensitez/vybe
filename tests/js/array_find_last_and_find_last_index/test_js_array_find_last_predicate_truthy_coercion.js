// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_predicate_truthy_coercion
// origin: languages/js/tests/js/test_js_array_find_last_and_find_last_index.rs

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

const arr = [0, 1, 0];
const found = arr.findLast(x => x); // Returns last truthy element
__check(__line(found), "1");
