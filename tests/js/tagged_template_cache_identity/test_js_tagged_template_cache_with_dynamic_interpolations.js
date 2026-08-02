// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_with_dynamic_interpolations
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

function tag(strings, ...values) { return strings; }
function getT(v) { return tag`Value: ${v}`; }
const t1 = getT("A");
const t2 = getT("B");
__check(__line(t1 === t2), "true"); // Same callsite -> identical template array reference!
