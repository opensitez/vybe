// vybe-test: js/structured_clone_map_set_regexp_date/test_js_structured_clone_regexp_named_groups_clone
// origin: languages/js/tests/js/test_js_structured_clone_map_set_regexp_date.rs

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

const re = /(?<year>\d{4})-(?<month>\d{2})/g;
const clone = structuredClone(re);
const match = clone.exec("2026-07");
__check(__line(match.groups.year + "|" + match.groups.month), "2026|07");
