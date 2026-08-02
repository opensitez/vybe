// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_arrow_function_lexical_brand_check
// origin: languages/js/tests/js/test_js_class_private_in_operator_brand_check.rs

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

class Container {
    #val = 50;
    getChecker() {
        return o => #val in o;
    }
}
const c = new Container();
const check = c.getChecker();
__check(__line(check(c) + "|" + check({})), "true|false");
