// vybe-test: js/regexp_named_capture_groups_indices/test_js_regexp_duplicate_named_capture_groups_in_disjunction_es2023
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

const re = /(?<date>\d{4}-\d{2})|(?<date>\d{2}\/\d{2})/; // Duplicate name across alternate branches
const m1 = re.exec("2026-07");
const m2 = re.exec("07/22");
__check(__line(m1.groups.date + "|" + m2.groups.date), "2026-07|07/22");
