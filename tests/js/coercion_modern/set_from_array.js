// vybe-test: js/coercion_modern/set_from_array
// origin: languages/js/tests/js/test_coercion_modern.rs

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

let s = new Set([1, 2, 2, 3, 3, 3]);
__check(__line(s.size), "3");
let arr = Array.from(s);
__check(__line(arr.join(",")), "1,2,3");
