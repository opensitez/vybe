// vybe-test: js/type_checking_patterns/array_check_patterns
// origin: languages/js/tests/js/test_type_checking_patterns.rs

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

const arr = [1, 2, 3];
__check(__line(Array.isArray(arr)), "true");
__check(__line(Array.isArray({})), "false");
__check(__line(Array.isArray("string")), "false");
__check(__line(arr instanceof Array), "true");
