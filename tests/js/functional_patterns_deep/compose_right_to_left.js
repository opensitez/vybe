// vybe-test: js/functional_patterns_deep/compose_right_to_left
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

const compose = (...fns) => x => fns.reduceRight((v, f) => f(v), x);
const add1 = x => x + 1;
const double = x => x * 2;
const square = x => x * x;
const transform = compose(add1, double, square);
// square(3) = 9, double(9) = 18, add1(18) = 19
__check(__line(transform(3)), "19");
