// vybe-test: js/tagged_template_cache_identity/test_js_tagged_template_cache_identity_constructor_method
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
class Item {
    constructor() {
        this.tpl = tag`CtorTemplate`;
    }
}
const i1 = new Item();
const i2 = new Item();
__check(__line(i1.tpl === i2.tpl), "true");
