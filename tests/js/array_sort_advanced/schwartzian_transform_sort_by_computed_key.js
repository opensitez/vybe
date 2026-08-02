// vybe-test: js/array_sort_advanced/schwartzian_transform_sort_by_computed_key
// origin: languages/js/tests/js/test_array_sort_advanced.rs

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

const words = ["banana", "fig", "cherry", "apple"];
const sorted = words
    .map(w => [w, w.length])
    .sort((a, b) => a[1] - b[1])
    .map(([w]) => w);
__check(__line(sorted.join(",")), "fig,apple,banana,cherry");
