// vybe-test: js/array_algorithms/frequency_map
// origin: languages/js/tests/js/test_array_algorithms.rs

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

function topK(arr, k) {
    const freq = new Map();
    for (const x of arr) freq.set(x, (freq.get(x) ?? 0) + 1);
    return [...freq.entries()]
        .sort((a, b) => b[1] - a[1])
        .slice(0, k)
        .map(([val]) => val);
}
const result = topK([1, 1, 1, 2, 2, 3], 2);
console.log(result[0]);
console.log(result[1]);
