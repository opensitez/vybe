// vybe-test: js/array_splice_to_spliced_slice_mutation/test_js_array_splice_frozen_array_throws_in_strict
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

const frozen = Object.freeze([1, 2, 3]);
try {
    "use strict";
    frozen.splice(0, 1);
} catch (e) {
    __check(__line("Splice Frozen Array TypeError"), "Splice Frozen Array TypeError");
}
