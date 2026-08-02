// vybe-test: js/functional_fp_patterns/either_monad
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

class Right {
    constructor(v) { this._v = v; }
    map(fn) { return new Right(fn(this._v)); }
    fold(_, f) { return f(this._v); }
}
class Left {
    constructor(v) { this._v = v; }
    map(_) { return this; }
    fold(f, _) { return f(this._v); }
}
const safe = (f, onError) => v => {
    try { return new Right(f(v)); }
    catch (e) { return new Left(onError(e)); }
};
const parseJSON = safe(JSON.parse, e => e.message);
const good = parseJSON('{"x":1}').map(o => o.x).fold(e => 0, v => v);
const bad = parseJSON("not json").fold(e => -1, v => v);
__check(__line(good), "1");
__check(__line(bad), "-1");
