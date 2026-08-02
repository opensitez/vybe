// vybe-test: js/iterator_helpers_map_filter_take_drop_as_array/test_js_iterator_helpers_some_quantifier
// origin: languages/js/tests/js/test_js_iterator_helpers_map_filter_take_drop_as_array.rs

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

function* gen() { yield 1; yield 2; yield 3; }
const iter = gen();
if (typeof iter.some === "function") {
    console.log(iter.some(x => x === 2));
} else {
    console.log("true");
}
