// vybe-test: js/class_patterns/string_prototype_user_method_resolves_off_primitive
// origin: languages/js/tests/js/test_class_patterns.rs

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

String.prototype.shout = function () { return this + "!"; };
__check(__line("hi".shout()), "hi!");
