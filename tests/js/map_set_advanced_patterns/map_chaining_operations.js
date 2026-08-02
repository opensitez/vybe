// vybe-test: js/map_set_advanced_patterns/map_chaining_operations
// origin: languages/js/tests/js/test_map_set_advanced_patterns.rs

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
const result = [...m.entries()]
    .filter(([, v]) => v > 1)
    .map(([k, v]) => k + "=" + v);
__check(__line(result.join(",")), "b=2,c=3");
