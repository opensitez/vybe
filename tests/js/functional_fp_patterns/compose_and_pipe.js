// vybe-test: js/functional_fp_patterns/compose_and_pipe
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

const compose = (...fns) => x => fns.reduceRight((v, f) => f(v), x);
const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const double = x => x * 2;
const addTen = x => x + 10;
const square = x => x * x;
const composed = compose(square, addTen, double);  // double -> addTen -> square
const piped = pipe(double, addTen, square);          // same
__check(__line(composed(5)), "400");  // (5*2+10)^2 = 400
__check(__line(piped(5)), "400");     // same
