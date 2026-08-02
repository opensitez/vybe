// vybe-test: js/regexp_v_flag_set_notation_intersection_subtraction/test_js_regexp_v_flag_escaped_special_chars_in_sets
// origin: languages/js/tests/js/test_js_regexp_v_flag_set_notation_intersection_subtraction.rs

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

const re = /[\(\)\[\]\{\}]/v;
__check(__line(re.test("(") + "|" + re.test("[") + "|" + re.test("a")), "true|true|false");
