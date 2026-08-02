// vybe-test: js/type_checking_patterns/duck_typing_check
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

function isIterable(val) {
    return val != null && typeof val[Symbol.iterator] === "function";
}
__check(__line(isIterable([1, 2, 3])), "true");
__check(__line(isIterable("string")), "true");
__check(__line(isIterable(new Map())), "true");
__check(__line(isIterable(42)), "false");
__check(__line(isIterable(null)), "false");
