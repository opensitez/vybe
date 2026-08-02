// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_index_basic_predicate
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

const arr = [5, 12, 50, 130, 44];
const idx = arr.findLastIndex(x => x > 45);
__check(__line(idx), "3");
