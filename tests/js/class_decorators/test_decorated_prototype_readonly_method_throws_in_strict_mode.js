// vybe-test: js/class_decorators/test_decorated_prototype_readonly_method_throws_in_strict_mode
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

"use strict";
class Target {
    action() { return "ok"; }
}
Object.defineProperty(Target.prototype, "action", { writable: false });
const t = new Target();
try {
    t.action = () => "changed";
} catch (e) {
    __check(__line(e.name), "TypeError");
}
