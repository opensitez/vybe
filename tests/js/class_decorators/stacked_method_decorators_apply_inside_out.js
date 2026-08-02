// vybe-test: js/class_decorators/stacked_method_decorators_apply_inside_out
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

function addA(fn) { return function(...args) { return fn.apply(this, args) + "A"; }; }
function addB(fn) { return function(...args) { return fn.apply(this, args) + "B"; }; }
class Str {
    hello() { return "X"; }
}
Str.prototype.hello = addA(addB(Str.prototype.hello));
__check(__line(new Str().hello()), "XBA");
