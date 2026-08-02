// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_raw_array_identity_matches_template
// origin: languages/js/tests/js/test_js_tagged_template_cache_identity.rs

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

function tag(strings) { return strings; }
function getRaw() { return tag`Text ${1}`.raw; }
const r1 = getRaw();
const r2 = getRaw();
__check(__line(r1 === r2), "true");
