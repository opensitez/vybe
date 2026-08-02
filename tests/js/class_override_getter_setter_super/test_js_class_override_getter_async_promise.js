// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_async_promise
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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
    get asyncVal() { return Promise.resolve(10); }
}
class Derived extends Base {
    get asyncVal() {
        return super.asyncVal.then(v => v * 4);
    }
}
new Derived().asyncVal.then(res => console.log(res));
