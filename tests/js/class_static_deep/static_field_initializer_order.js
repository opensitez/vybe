// vybe-test: js/class_static_deep/static_field_initializer_order
// origin: languages/js/tests/js/test_class_static_deep.rs

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

const log = [];
class Foo {
    static a = (log.push("a"), 1);
    static b = (log.push("b"), 2);
    static c = (log.push("c"), 3);
}
__check(__line(log.join(",")), "a,b,c");
