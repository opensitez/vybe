// vybe-test: js/string_array_advanced/array_sort_stable
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

let items = [
    { name: "A", score: 1 },
    { name: "B", score: 2 },
    { name: "C", score: 1 },
    { name: "D", score: 2 }
];
items.sort((a, b) => a.score - b.score);
__check(__line(items.map(i => i.name).join(",")), "A,C,B,D");
