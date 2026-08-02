// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_index_returns_minus_one_when_not_found
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

const arr = [1, 2, 3];
const idx = arr.findLastIndex(x => x > 10);
__check(__line(idx), "-1");
