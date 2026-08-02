// vybe-test: js/mixin_abstract_patterns/abstract_class_pattern_with_new_target
// origin: languages/js/tests/js/test_mixin_abstract_patterns.rs

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

class Abstract {
    constructor() {
        if (new.target === Abstract) throw new Error("Cannot instantiate Abstract");
    }
    method() { throw new Error("Not implemented"); }
}
class Concrete extends Abstract {
    method() { return "implemented"; }
}
let threw = false;
try { new Abstract(); } catch { threw = true; }
__check(__line(threw), "true");
__check(__line(new Concrete().method()), "implemented");
