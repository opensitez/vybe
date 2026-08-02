// vybe-test: js/regexp_multiline_m_and_dotall_s_flags/test_js_regexp_constructor_flag_override
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

const re1 = /a/g;
const re2 = new RegExp(re1, "i"); // Passing flags to RegExp constructor overrides original pattern flags!
__check(__line(re2.flags), "i");
