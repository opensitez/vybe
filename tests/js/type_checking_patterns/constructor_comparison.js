// vybe-test: js/type_checking_patterns/constructor_comparison
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

const arr = [];
const obj = {};
const fn = function() {};
__check(__line(arr.constructor === Array), "true");
__check(__line(obj.constructor === Object), "true");
__check(__line(fn.constructor === Function), "true");
