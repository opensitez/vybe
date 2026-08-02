// vybe-test: js/functional_patterns_deep/functor_map_over_custom_type
// origin: languages/js/tests/js/test_functional_patterns_deep.rs

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
    constructor(val) { this.val = val; }
    static of(val) { return new Maybe(val); }
    map(fn) {
        return this.val == null ? this : Maybe.of(fn(this.val));
    }
    get() { return this.val; }
}
const result = Maybe.of(5)
    .map(x => x * 2)
    .map(x => x + 1)
    .get();
__check(__line(result), "11");
const nullResult = Maybe.of(null).map(x => x * 2).get();
__check(__line(nullResult), "null");
