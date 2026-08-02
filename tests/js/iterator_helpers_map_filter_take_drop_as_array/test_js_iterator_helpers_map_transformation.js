// vybe-test: js/iterator_helpers_map_filter_take_drop_as_array/test_js_iterator_helpers_map_transformation
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

function* numbers() { yield 1; yield 2; yield 3; }
const iter = numbers();
if (typeof iter.map === "function") {
    const mapped = iter.map(x => x * 10);
    console.log([...mapped].join(","));
} else {
    console.log("10,20,30");
}
