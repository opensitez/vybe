// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_replaceall_with_named_capture_groups
// origin: languages/js/tests/js/test_js_regexp_string_match_all_replace_all.rs

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

const str = "2026-07-22";
const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/g;
const res = str.replaceAll(re, "$<month>/$<day>/$<year>");
__check(__line(res), "07/22/2026");
