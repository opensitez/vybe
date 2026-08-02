// vybe-test: js/array_algorithms/quicksort
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

function quicksort(arr) {
    if (arr.length <= 1) return arr;
    const pivot = arr[arr.length >> 1];
    const left = arr.filter(x => x < pivot);
    const mid = arr.filter(x => x === pivot);
    const right = arr.filter(x => x > pivot);
    return [...quicksort(left), ...mid, ...quicksort(right)];
}
__check(__line(quicksort([3, 1, 4, 1, 5, 9, 2, 6]).join(",")), "1,1,2,3,4,5,6,9");
