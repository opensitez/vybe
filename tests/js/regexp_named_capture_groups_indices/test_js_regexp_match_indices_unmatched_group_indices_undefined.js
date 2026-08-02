// vybe-test: js/regexp_named_capture_groups_indices/test_js_regexp_match_indices_unmatched_group_indices_undefined
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

const re = /(?<x>a)|(?<y>b)/d;
const match = re.exec("a");
__check(__line(match.indices[1].join(":") + "|" + (match.indices[2] === undefined)), "0:1|true");
