// vybe-test: js/closure_scope_deep/compose_right_to_left
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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
const double = x => x * 2;
const addOne = x => x + 1;
const square = x => x * x;

const transform = compose(double, addOne, square);
// square(3) = 9, addOne(9) = 10, double(10) = 20
__check(__line(transform(3)), "20");
