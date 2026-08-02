// vybe-test: js/reflect_apply_construct_get_set_methods/test_js_reflect_own_keys_strings_and_symbols
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

const sym = Symbol("s");
const obj = { b: 2, a: 1, [sym]: 3 };
const keys = Reflect.ownKeys(obj);
__check(__line(keys.length + "|" + (keys[2] === sym)), "3|true");
