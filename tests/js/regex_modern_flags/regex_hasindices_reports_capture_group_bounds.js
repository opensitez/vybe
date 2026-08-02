// vybe-test: js/regex_modern_flags/regex_hasindices_reports_capture_group_bounds
// origin: languages/js/tests/js/test_regex_modern_flags.rs

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

const re = /(\d+)-(\d+)/d;
const match = re.exec("date 2024-07");
__check(__line(match.indices[1].join(",")), "5,9");
__check(__line(match.indices[2].join(",")), "10,12");
