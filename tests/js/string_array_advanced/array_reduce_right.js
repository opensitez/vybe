// vybe-test: js/string_array_advanced/array_reduce_right
// origin: languages/js/tests/js/test_string_array_advanced.rs

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

let arr = ["a", "b", "c", "d"];
let result = arr.reduceRight((acc, val) => acc + val, "");
__check(__line(result), "dcba");
