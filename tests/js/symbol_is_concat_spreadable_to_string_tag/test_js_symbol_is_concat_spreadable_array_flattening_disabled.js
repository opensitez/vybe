// vybe-test: js/symbol_is_concat_spreadable_to_string_tag/test_js_symbol_is_concat_spreadable_array_flattening_disabled
// origin: languages/js/tests/js/test_js_symbol_is_concat_spreadable_to_string_tag.rs

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

const arr1 = [1, 2];
const arr2 = [3, 4];
arr2[Symbol.isConcatSpreadable] = false; // Prevents concat from flattening arr2!

const res = arr1.concat(arr2);
__check(__line(res.length + "|" + Array.isArray(res[2])), "3|true");
