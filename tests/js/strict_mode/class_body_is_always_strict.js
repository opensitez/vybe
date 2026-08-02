// vybe-test: js/strict_mode/class_body_is_always_strict
// origin: languages/js/tests/js/test_strict_mode.rs

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

class Foo {
    method() {
        return this === undefined ? "strict" : "sloppy";
    }
}
const fn2 = Foo.prototype.method;
__check(__line(fn2()), "strict");
