// vybe-test: js/regexp_multiline_m_and_dotall_s_flags/test_js_regexp_has_indices_d_flag_property
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

const re = /a/d;
__check(__line(re.hasIndices + "|" + re.flags), "true|d");
