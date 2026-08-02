// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_nested_tagged_templates
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
function getNested() {
    return tag`Outer ${tag`Inner`}`;
}
const n1 = getNested();
const n2 = getNested();
__check(__line(n1 === n2), "true");
