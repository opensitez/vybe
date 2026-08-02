// vybe-test: js/regexp_named_capture_groups_indices/test_js_regexp_unicode_identifier_named_groups
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

const re = /(?<π>\d+)/u;
const match = re.exec("314");
__check(__line(match.groups.π), "314");
