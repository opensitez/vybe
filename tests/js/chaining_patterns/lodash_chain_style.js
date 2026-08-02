// vybe-test: js/chaining_patterns/lodash_chain_style
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

// Simplified lodash-like chain
class Chain {
    constructor(val) { this._val = val; }
    map(fn) { return new Chain(this._val.map(fn)); }
    filter(fn) { return new Chain(this._val.filter(fn)); }
    reduce(fn, init) { return this._val.reduce(fn, init); }
    value() { return this._val; }
}
const result = new Chain([1, 2, 3, 4, 5, 6])
    .filter(x => x % 2 === 0)
    .map(x => x * x)
    .reduce((a, b) => a + b, 0);
__check(__line(result), "56"); // 4+16+36 = 56
