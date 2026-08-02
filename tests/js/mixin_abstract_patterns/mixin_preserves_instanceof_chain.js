// vybe-test: js/mixin_abstract_patterns/mixin_preserves_instanceof_chain
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

const Mixin = Base => class extends Base {};
class Root {}
class Child extends Mixin(Root) {}
const c = new Child();
__check(__line(c instanceof Child), "true");
__check(__line(c instanceof Root), "true");
