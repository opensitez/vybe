// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_identity_closure_rebinding
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

function makeTagger() {
    const tag = strings => strings;
    return () => tag`ClosureTemplate`;
}
const f1 = makeTagger();
const f2 = makeTagger();
__check(__line(f1() === f1()), "true");
