// vybe-test: js/reflect_apply_construct_get_set/test_js_reflect_construct_non_constructor_target_throws_typeerror
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set.rs

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

try {
    Reflect.construct(() => {}, []);
} catch (e) {
    __check(__line("Reflect.construct Non-Constructor TypeError"), "Reflect.construct Non-Constructor TypeError");
}
