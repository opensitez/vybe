// vybe-test: js/array_algorithms/merge_sort
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

function merge(a, b) {
    const result = [];
    let i = 0, j = 0;
    while (i < a.length && j < b.length) {
        result.push(a[i] <= b[j] ? a[i++] : b[j++]);
    }
    return [...result, ...a.slice(i), ...b.slice(j)];
}
function mergeSort(arr) {
    if (arr.length <= 1) return arr;
    const mid = arr.length >> 1;
    return merge(mergeSort(arr.slice(0, mid)), mergeSort(arr.slice(mid)));
}
console.log(mergeSort([5, 3, 8, 1, 9, 2]).join(","));
