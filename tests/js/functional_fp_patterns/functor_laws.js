// vybe-test: js/functional_fp_patterns/functor_laws
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

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

class Box {
    constructor(v) { this._v = v; }
    map(fn) { return new Box(fn(this._v)); }
    value() { return this._v; }
}
const identity = x => x;
const double = x => x * 2;
const addOne = x => x + 1;
// Identity law: map(id) === id
__check(__line(new Box(5).map(identity).value() === new Box(5).value()), "true");
// Composition law: map(f).map(g) === map(g(f(x)))
const a = new Box(5).map(double).map(addOne).value();
const b = new Box(5).map(x => addOne(double(x))).value();
__check(__line(a === b), "true");
