// vybe-test: js/immutable_patterns/immutable_array_remove
// origin: languages/js/tests/js/test_immutable_patterns.rs

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

const arr = [1, 2, 3, 4, 5];
function removeAt(arr, index) {
    return [...arr.slice(0, index), ...arr.slice(index + 1)];
}
const result = removeAt(arr, 2);
__check(__line(arr.join(",")), "1,2,3,4,5");    // original unchanged
__check(__line(result.join(",")), "1,2,4,5");
