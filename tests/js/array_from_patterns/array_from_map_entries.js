// vybe-test: js/array_from_patterns/array_from_map_entries
// origin: languages/js/tests/js/test_array_from_patterns.rs

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

const m = new Map([["a", 1], ["b", 2]]);
const arr = Array.from(m);
__check(__line(arr.map(([k, v]) => k + "=" + v).join(",")), "a=1,b=2");
