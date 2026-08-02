// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_in_loop_filtering
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

class Target {
    #id = 1;
    static countValid(arr) {
        return arr.filter(item => #id in item).length;
    }
}
const items = [new Target(), {}, new Target(), "str"];
__check(__line(Target.countValid(items)), "2");
