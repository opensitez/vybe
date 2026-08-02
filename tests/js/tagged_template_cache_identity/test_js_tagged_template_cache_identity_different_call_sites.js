// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_identity_different_call_sites
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
const t1 = tag`Same Text`;
const t2 = tag`Same Text`;
__check(__line(t1 === t2), "false"); // Per ES2018 spec, different call sites produce distinct template objects
