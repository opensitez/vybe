// vybe-test: js/regex_string_methods/replace_with_capture_group_reference
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

const result = "2024-01-15".replace(/(\d{4})-(\d{2})-(\d{2})/, "$3/$2/$1");
__check(__line(result), "15/01/2024");
