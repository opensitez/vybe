// vybe-test: js/array_sort_advanced/sort_stable_complex_key
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

const data = [
    { key: "b", order: 0 },
    { key: "a", order: 1 },
    { key: "b", order: 2 },
    { key: "a", order: 3 },
];
data.sort((a, b) => a.key.localeCompare(b.key));
// Stable: a's keep order 1,3; b's keep order 0,2
__check(__line(data[0].key + data[0].order), "a1");
__check(__line(data[1].key + data[1].order), "a3");
__check(__line(data[2].key + data[2].order), "b0");
__check(__line(data[3].key + data[3].order), "b2");
