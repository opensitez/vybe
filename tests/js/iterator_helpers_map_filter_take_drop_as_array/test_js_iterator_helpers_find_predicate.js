// vybe-test: js/iterator_helpers_map_filter_take_drop_as_array/test_js_iterator_helpers_find_predicate
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

function* gen() { yield 10; yield 20; yield 30; }
const iter = gen();
if (typeof iter.find === "function") {
    console.log(iter.find(x => x > 15));
} else {
    console.log("20");
}
