// vybe-test: js/array_algorithms/run_length_encoding
// origin: languages/js/tests/js/test_array_algorithms.rs

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

function rle(str) {
    return str.replace(/(.)\1*/g, (m, c) => m.length > 1 ? m.length + c : c);
}
__check(__line(rle("aabbbccddddee")), "2a3b2c4d2e");
__check(__line(rle("abc")), "abc");
__check(__line(rle("aaaa")), "4a");
