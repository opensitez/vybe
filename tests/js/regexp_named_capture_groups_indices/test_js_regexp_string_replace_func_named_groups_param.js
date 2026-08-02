// vybe-test: js/regexp_named_capture_groups_indices/test_js_regexp_string_replace_func_named_groups_param
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

const re = /(?<num>\d+)/;
const res = "Item 42".replace(re, (match, p1, offset, string, groups) => {
    return String(Number(groups.num) * 2);
});
__check(__line(res), "Item 84");
