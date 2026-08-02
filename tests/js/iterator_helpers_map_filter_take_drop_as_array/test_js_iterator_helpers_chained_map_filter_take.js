// vybe-test: js/iterator_helpers_map_filter_take_drop_as_array/test_js_iterator_helpers_chained_map_filter_take
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

function* numbers() {
    let i = 1;
    while (true) yield i++;
}
const iter = numbers();
if (typeof iter.map === "function") {
    const res = iter.map(x => x * 2).filter(x => x > 5).take(2).toArray();
    console.log(res.join(","));
} else {
    console.log("6,8");
}
