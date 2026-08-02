// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_index_array_like_object
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

const arrayLike = { 0: "first", 1: "second", length: 2 };
const idx = Array.prototype.findLastIndex.call(arrayLike, x => x.startsWith("f"));
__check(__line(idx), "0");
