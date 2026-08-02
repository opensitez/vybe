// vybe-test: js/functional_patterns_deep/pipe_left_to_right
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

const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const transform = pipe(
    x => x * 2,
    x => x + 1,
    x => x.toString()
);
__check(__line(transform(5)), "11"); // 5*2=10, +1=11, toString="11"
