// vybe-test: js/regexp_multiline_m_and_dotall_s_flags/test_js_regexp_global_g_flag_replace_all_matches
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

const str = "foo foo foo";
__check(__line(str.replace(/foo/g, "bar") + "|" + str.replace(/foo/, "bar")), "bar bar bar|bar foo foo");
