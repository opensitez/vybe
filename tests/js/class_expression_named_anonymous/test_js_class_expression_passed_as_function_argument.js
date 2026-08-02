// vybe-test: js/class_expression_named_anonymous/test_js_class_expression_passed_as_function_argument
// origin: languages/js/tests/js/test_js_class_expression_named_anonymous.rs

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

function instantiate(ClassRef, arg) {
    return new ClassRef(arg);
}
const item = instantiate(class {
    constructor(v) { this.v = v; }
}, "ArgumentVal");
__check(__line(item.v), "ArgumentVal");
