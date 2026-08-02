// vybe-test: js/array_sort_advanced/sort_is_stable
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

// Stable sort preserves relative order of equal elements
const items = [
    { priority: 1, name: "c" },
    { priority: 2, name: "a" },
    { priority: 1, name: "b" },
];
items.sort((a, b) => a.priority - b.priority);
// Equal priority items (c and b) must retain c before b
__check(__line(items[0].name), "c");
__check(__line(items[1].name), "b");
