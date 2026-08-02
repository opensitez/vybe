// vybe-test: js/array_sort_advanced/sort_mixed_types_coerced_to_string
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

const arr = [null, 1, "a"];
arr.sort();
// null → "null", 1 → "1", "a" → "a"
// lexicographic: "1" < "a" < "null"
__check(__line(arr[0]), "1");
