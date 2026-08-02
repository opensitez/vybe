// vybe-test: js/array_sort_to_sorted_comparator_stability/test_js_array_sort_stability_es2019
// origin: languages/js/tests/js/test_js_array_sort_to_sorted_comparator_stability.rs

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

const items = [
    { name: "A", score: 10 },
    { name: "B", score: 5 },
    { name: "C", score: 10 },
    { name: "D", score: 5 }
];
items.sort((a, b) => a.score - b.score);
__check(__line(items.map(i => i.name).join(",")), "B,D,A,C"); // Stable sort preserves relative order: B, D, A, C
