// vybe-test: js/array_from_patterns/array_from_length_property_fills_undefined
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

const arr = Array.from({ length: 3 });
__check(__line(arr.length), "3");
__check(__line(arr.every(x => x === undefined)), "true");
