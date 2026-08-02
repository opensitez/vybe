// vybe-test: js/string_methods_deep/slice_beyond_length
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

const s = "hi";
__check(__line(s.slice(0, 100)), "hi");
__check(__line(s.slice(5)), "");
