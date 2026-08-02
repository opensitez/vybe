// vybe-test: js/array_find_last_and_find_last_index/test_js_typed_array_find_last
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

const u8 = new Uint8Array([10, 20, 30, 40]);
const found = u8.findLast(x => x < 35);
__check(__line(found), "30");
