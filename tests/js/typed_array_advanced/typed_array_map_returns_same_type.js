// vybe-test: js/typed_array_advanced/typed_array_map_returns_same_type
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const ta = new Int32Array([1, 2, 3]);
const mapped = ta.map(x => x * 2);
__check(__line(mapped.length === 3), "true");
__check(__line(Array.from(mapped).join(",")), "2,4,6");
