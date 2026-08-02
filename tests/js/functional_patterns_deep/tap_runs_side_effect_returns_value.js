// vybe-test: js/functional_patterns_deep/tap_runs_side_effect_returns_value
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

const tap = (fn) => (x) => { fn(x); return x; };
const log = tap(x => console.log("tap: " + x));
const result = [1, 2, 3].map(log);
console.log(result.join(","));
