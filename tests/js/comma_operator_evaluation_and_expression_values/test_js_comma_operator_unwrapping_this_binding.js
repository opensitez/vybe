// vybe-test: js/comma_operator_evaluation_and_expression_values/test_js_comma_operator_unwrapping_this_binding
// origin: languages/js/tests/js/test_js_comma_operator_evaluation_and_expression_values.rs

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

const obj = {
    val: "ObjVal",
    getVal() { return this ? this.val : "NoThis"; }
};
__check(__line(obj.getVal() + "|" + (0, obj.getVal)()), "ObjVal|NoThis"); // (0, obj.getVal)() invokes method with this = undefined/globalThis!
