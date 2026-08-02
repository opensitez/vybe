// vybe-test: js/chaining_patterns/functional_chain_point_free
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

const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const process = pipe(
    arr => arr.filter(x => x > 2),
    arr => arr.map(x => x * 2),
    arr => arr.reduce((a, b) => a + b, 0)
);
__check(__line(process([1, 2, 3, 4, 5])), "24"); // [3,4,5] -> [6,8,10] -> 24
