// vybe-test: js/regex_v_flag/v_flag_nested_class_union
// origin: languages/js/tests/js/test_regex_v_flag.rs

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

// Hex digits: 0-9 or a-f or A-F
const re = /^[0-9a-fA-F]+$/;
__check(__line(re.test("deadbeef123")), "true");
__check(__line(re.test("xyz")), "false");
