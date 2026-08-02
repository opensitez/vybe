// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_in_class_method
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

class Parser {
    tag(strings) { return strings; }
    parse() { return this.tag`ClassTemplate`; }
}
const p = new Parser();
__check(__line(p.parse() === p.parse()), "true");
