// vybe-test: js/class_super_method_property_access/test_js_class_super_method_async_await
// origin: languages/js/tests/js/test_js_class_super_method_property_access.rs

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

class Base {
    async load() { return await Promise.resolve("BaseLoad"); }
}
class Sub extends Base {
    async load() {
        const val = await super.load();
        return val + "_Extended";
    }
}
new Sub().load().then(res => console.log(res));
