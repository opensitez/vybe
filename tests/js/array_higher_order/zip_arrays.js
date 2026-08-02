// vybe-test: js/array_higher_order/zip_arrays
// origin: languages/js/tests/js/test_array_higher_order.rs

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

function zip(a, b) {
    const len = Math.min(a.length, b.length);
    const out = [];
    for (let i = 0; i < len; i++) {
        out.push([a[i], b[i]]);
    }
    return out;
}
const zipped = zip([1, 2, 3], ["a", "b", "c"], [true, false, true]);
console.log(zipped[0].join(","));
console.log(zipped[1].join(","));
