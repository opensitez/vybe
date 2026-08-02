// vybe-test: js/class_decorators/class_decorator_context_kind_is_class
// origin: languages/js/tests/js/test_class_decorators.rs

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

let capturedKind;
function capture(target, kind) { capturedKind = kind; }
class Foo {}
capture(Foo, "class");
__check(__line(capturedKind), "class");
