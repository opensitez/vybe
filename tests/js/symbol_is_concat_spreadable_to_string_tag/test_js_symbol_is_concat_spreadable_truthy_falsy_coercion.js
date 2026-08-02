// vybe-test: js/symbol_is_concat_spreadable_to_string_tag/test_js_symbol_is_concat_spreadable_truthy_falsy_coercion
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

const arr = [10, 20];
arr[Symbol.isConcatSpreadable] = 0; // Falsy -> not spreadable
const res1 = [0].concat(arr);

arr[Symbol.isConcatSpreadable] = 1; // Truthy -> spreadable
const res2 = [0].concat(arr);
__check(__line(res1.length + "|" + res2.length), "2|3");
