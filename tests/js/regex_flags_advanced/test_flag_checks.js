// vybe-test: js/regex_flags_advanced/test_flag_checks
// origin: languages/js/tests/js/test_regex_flags_advanced.rs

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

const re = /abc/gim;
__check(__line(re.global), "true");
__check(__line(re.ignoreCase), "true");
__check(__line(re.multiline), "true");
__check(__line(re.flags.split("").sort().join("")), "gim");
