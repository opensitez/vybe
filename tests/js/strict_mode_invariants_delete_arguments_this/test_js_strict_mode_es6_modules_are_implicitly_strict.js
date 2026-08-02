// vybe-test: js/strict_mode_invariants_delete_arguments_this/test_js_strict_mode_es6_modules_are_implicitly_strict
// origin: languages/js/tests/js/test_js_strict_mode_invariants_delete_arguments_this.rs

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

function fnInModule() {
    try {
        eval("undeclaredInModule = 10;");
    } catch (e) {
        return true;
    }
    return false;
}
__check(__line(fnInModule()), "true");
