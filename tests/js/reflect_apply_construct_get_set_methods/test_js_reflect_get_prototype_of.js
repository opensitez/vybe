// vybe-test: js/reflect_apply_construct_get_set_methods/test_js_reflect_get_prototype_of
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

const proto = { val: 1 };
const obj = Object.create(proto);
__check(__line(Reflect.getPrototypeOf(obj) === proto), "true");
