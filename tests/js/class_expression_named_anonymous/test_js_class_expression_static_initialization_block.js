// vybe-test: js/class_expression_named_anonymous/test_js_class_expression_static_initialization_block
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

const App = class {
    static status;
    static {
        this.status = "Ready";
    }
};
__check(__line(App.status), "Ready");
