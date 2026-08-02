// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_in_generator_function
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
function* gen() {
    yield tag`GenTemplate`;
    yield tag`GenTemplate`;
}
const g = gen();
const t1 = g.next().value;
const t2 = g.next().value;
__check(__line(t1 === t2), "false"); // Different call sites inside generator yield distinct objects
