// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_identity_same_call_site
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

function tag(strings) {
    return strings;
}
function getTemplate() {
    return tag`Hello ${1}`;
}
const t1 = getTemplate();
const t2 = getTemplate();
__check(__line(t1 === t2), "true");
