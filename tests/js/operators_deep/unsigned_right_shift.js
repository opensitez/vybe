// vybe-test: js/operators_deep/unsigned_right_shift
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line(-1 >>> 0), "4294967295");  // 4294967295 (convert to uint32)
__check(__line(0xFFFFFFFF >>> 0), "4294967295"); // 4294967295
