// vybe-test: js/array_from_patterns/array_of_vs_array_from
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

// Array.of: treats args as elements
const a = Array.of(7);
__check(__line(a.length), "1");
__check(__line(a[0]), "7");
// Array(7): creates hole array of length 7
const b = Array(7);
__check(__line(b.length), "7");
__check(__line(b[0]), "undefined");
