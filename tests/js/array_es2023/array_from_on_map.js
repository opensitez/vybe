// vybe-test: js/array_es2023/array_from_on_map
// origin: languages/js/tests/js/test_array_es2023.rs

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

const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
const arr = Array.from(m);
__check(__line(arr.length), "3");
__check(__line(arr[0][0] + ":" + arr[0][1]), "a:1");
