// vybe-test: js/array_from_patterns/array_from_length_with_map
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

const arr = Array.from({ length: 5 }, (_, i) => i * i);
__check(__line(arr.join(",")), "0,1,4,9,16");
