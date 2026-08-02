// vybe-test: js/reflect_apply_construct_get_set_methods/test_js_reflect_set_non_writable_returns_false
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
Object.defineProperty(obj, "fixed", { value: 10, writable: false });
const success = Reflect.set(obj, "fixed", 99);
__check(__line(success + "|" + obj.fixed), "false|10");
