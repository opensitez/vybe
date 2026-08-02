// vybe-test: js/chaining_patterns/promise_style_chaining
// origin: languages/js/tests/js/test_chaining_patterns.rs

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

class Computation {
    constructor(val) { this._val = val; }
    map(fn) { return new Computation(fn(this._val)); }
    flatMap(fn) { return fn(this._val); }
    getOrElse(def) { return this._val ?? def; }
}
const result = new Computation(5)
    .map(x => x * 2)
    .map(x => x + 1)
    .flatMap(x => new Computation(x.toString()))
    .getOrElse("default");
__check(__line(result), "11");
