// vybe-test: js/iterator_helpers_map_filter_take_drop_as_array/test_js_iterator_helpers_lazy_evaluation
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

let evaluatedCount = 0;
function* gen() {
    while (true) {
        evaluatedCount++;
        yield evaluatedCount;
    }
}
const iter = gen();
if (typeof iter.map === "function") {
    const mapped = iter.map(x => x);
    console.log(evaluatedCount); // Generator has NOT evaluated any yields yet!
} else {
    console.log("0");
}
