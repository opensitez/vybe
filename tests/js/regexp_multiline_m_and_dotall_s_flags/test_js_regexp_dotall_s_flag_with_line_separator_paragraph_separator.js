// vybe-test: js/regexp_multiline_m_and_dotall_s_flags/test_js_regexp_dotall_s_flag_with_line_separator_paragraph_separator
// origin: languages/js/tests/js/test_js_regexp_multiline_m_and_dotall_s_flags.rs

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

const str = "a\u2028b\u2029c"; // Line Separator U+2028 & Paragraph Separator U+2029
__check(__line(str.match(/a.*c/s) !== null), "true");
