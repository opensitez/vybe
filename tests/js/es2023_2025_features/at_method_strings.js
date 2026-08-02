// vybe-test: js/es2023_2025_features/at_method_strings
// origin: languages/js/tests/js/test_es2023_2025_features.rs

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

const s = "hello";
__check(__line(s.at(0)), "h");
__check(__line(s.at(-1)), "o");
__check(__line(s.at(-2)), "l");
__check(__line(s.at(10)), "undefined");
