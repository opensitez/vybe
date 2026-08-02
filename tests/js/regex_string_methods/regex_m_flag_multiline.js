// vybe-test: js/regex_string_methods/regex_m_flag_multiline
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

const re = /^\d+/mg;
const text = "1 hello\n2 world\n3 foo";
const matches = text.match(re);
__check(__line(matches.join(",")), "1,2,3");
