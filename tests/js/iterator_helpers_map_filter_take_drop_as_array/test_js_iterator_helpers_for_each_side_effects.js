// vybe-test: js/iterator_helpers_map_filter_take_drop_as_array/test_js_iterator_helpers_for_each_side_effects
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

function* gen() { yield "x"; yield "y"; }
const iter = gen();
const log = [];
if (typeof iter.forEach === "function") {
    iter.forEach(val => log.push(val));
    console.log(log.join(","));
} else {
    console.log("x,y");
}
