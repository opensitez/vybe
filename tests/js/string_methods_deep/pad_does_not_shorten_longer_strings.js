// vybe-test: js/string_methods_deep/pad_does_not_shorten_longer_strings
// origin: languages/js/tests/js/test_string_methods_deep.rs

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
__check(__line(s.padStart(3)), "hello");
__check(__line(s.padEnd(3)), "hello");
