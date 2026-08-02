// vybe-test: js/functional_fp_patterns/maybe_monad
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

class Maybe {
    constructor(v) { this._v = v; }
    static of(v) { return new Maybe(v); }
    isNothing() { return this._v == null; }
    map(fn) { return this.isNothing() ? this : Maybe.of(fn(this._v)); }
    getOrElse(def) { return this.isNothing() ? def : this._v; }
}
const result1 = Maybe.of(5).map(x => x * 2).map(x => x + 1).getOrElse(0);
const result2 = Maybe.of(null).map(x => x * 2).getOrElse(-1);
__check(__line(result1), "11");
__check(__line(result2), "-1");
