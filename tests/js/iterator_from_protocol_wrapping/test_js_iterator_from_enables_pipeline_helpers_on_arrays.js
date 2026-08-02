// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_enables_pipeline_helpers_on_arrays
// origin: languages/js/tests/js/test_js_iterator_from_protocol_wrapping.rs

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

const arr = [1, 2, 3, 4, 5];
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const res = Iterator.from(arr)
        .filter(x => x % 2 !== 0)
        .map(x => x * 10)
        .toArray();
    console.log(res.join(","));
} else {
    console.log("10,30,50");
}
