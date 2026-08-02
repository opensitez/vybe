// vybe-test: js/string_pad_start_pad_end_repeat_methods/test_js_string_pad_start_default_pad_string_is_space
// origin: languages/js/tests/js/test_js_string_pad_start_pad_end_repeat_methods.rs

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

const str = "abc";
__check(__line("'" + str.padStart(5) + "'"), "'  abc'");
