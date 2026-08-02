// vybe-test: js/regexp_named_capture_groups_indices/test_js_regexp_named_capture_groups_basic
// origin: languages/js/tests/js/test_js_regexp_named_capture_groups_indices.rs

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

const re = /(?<year>\d{4})-(?<month>\d{2})-(?<day>\d{2})/;
const match = re.exec("2026-07-22");
__check(__line(`${match.groups.year}:${match.groups.month}:${match.groups.day}`), "2026:07:22");
