// vybe-test: js/unary_plus_minus_tilde_void_typeof_delete/test_js_delete_non_configurable_property_returns_false_in_non_strict
// origin: languages/js/tests/js/test_js_unary_plus_minus_tilde_void_typeof_delete.rs

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
Object.defineProperty(obj, "fixed", { value: 1, configurable: false });
const res = delete obj.fixed;
__check(__line(res + "|fixed=" + obj.fixed), "false|fixed=1");
