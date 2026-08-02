// vybe-test: js/class_expression_named_anonymous/test_js_named_class_expression_internal_name_binding
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

const FactorialClass = class Fact {
    static calc(n) {
        if (n <= 1) return 1;
        return n * Fact.calc(n - 1); // Fact internal binding accessible inside class body!
    }
};
__check(__line(FactorialClass.calc(5)), "120");
