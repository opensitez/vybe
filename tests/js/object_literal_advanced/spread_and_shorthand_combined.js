// vybe-test: js/object_literal_advanced/spread_and_shorthand_combined
// origin: languages/js/tests/js/test_object_literal_advanced.rs

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

const base = { a: 1, b: 2 };
const c = 3;
const result = { ...base, c, d: 4 };
__check(__line(result.a), "1");
__check(__line(result.c), "3");
__check(__line(result.d), "4");
