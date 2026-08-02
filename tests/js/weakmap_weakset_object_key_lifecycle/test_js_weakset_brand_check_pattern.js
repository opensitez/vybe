// vybe-test: js/weakmap_weakset_object_key_lifecycle/test_js_weakset_brand_check_pattern
// origin: languages/js/tests/js/test_js_weakmap_weakset_object_key_lifecycle.rs

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

const brandStore = new WeakSet();
class CustomBrand {
    constructor() {
        brandStore.add(this);
    }
    static isInstance(obj) {
        return brandStore.has(obj);
    }
}
const b = new CustomBrand();
__check(__line(CustomBrand.isInstance(b) + "|" + CustomBrand.isInstance({})), "true|false");
