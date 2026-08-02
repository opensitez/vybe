// vybe-test: js/reflect_apply_construct_get_set_methods/test_js_reflect_set_prototype_of_returns_boolean
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set_methods.rs

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

const obj = {};
const proto = { newProto: true };
const success = Reflect.setPrototypeOf(obj, proto);
__check(__line(success + "|" + obj.newProto), "true|true");
