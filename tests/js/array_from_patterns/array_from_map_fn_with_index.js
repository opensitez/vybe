// vybe-test: js/array_from_patterns/array_from_map_fn_with_index
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

const arr = Array.from("abc", (c, i) => i + ":" + c);
__check(__line(arr.join(",")), "0:a,1:b,2:c");
