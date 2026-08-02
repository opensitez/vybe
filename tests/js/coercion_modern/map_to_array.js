// vybe-test: js/coercion_modern/map_to_array
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

let m = new Map([["x", 10], ["y", 20]]);
let arr = Array.from(m);
__check(__line(arr.length), "2");
__check(__line(arr[0][0] + "=" + arr[0][1]), "x=10");
