// vybe-test: js/class_expression_named_anonymous/test_js_named_class_expression_internal_name_immutable
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

const Foo = class Bar {
    static tryRebind() {
        "use strict";
        try {
            Bar = 123; // Internal class expression binding is read-only constant!
        } catch (e) {
            __check(__line("Internal Name Immutable TypeError"), "Internal Name Immutable TypeError");
        }
    }
};
Foo.tryRebind();
