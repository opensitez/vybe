// vybe-test: js/regex_v_flag/v_flag_creates_valid_regex
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

const re = /[abc]/u;
__check(__line(re.flags.includes("u")), "true");
__check(__line(re.test("a")), "true");
__check(__line(re.test("d")), "false");
