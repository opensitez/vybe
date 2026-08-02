// vybe-test: js/class_expression_named_anonymous/test_js_class_expression_in_object_literal_property
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

const factory = {
    UserClass: class {
        constructor(name) { this.name = name; }
    }
};
const u = new factory.UserClass("Bob");
__check(__line(u.name), "Bob");
