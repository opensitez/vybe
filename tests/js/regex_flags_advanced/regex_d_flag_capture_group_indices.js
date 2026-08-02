// vybe-test: js/regex_flags_advanced/regex_d_flag_capture_group_indices
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

const re = /(\d{4})-(\d{2})/d;
const m = re.exec("Date: 2024-06");
__check(__line(m.indices[1].join(",")), "6,10"); // year capture indices
__check(__line(m.indices[2].join(",")), "11,13"); // month capture
