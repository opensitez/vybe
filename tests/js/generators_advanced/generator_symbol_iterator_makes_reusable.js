// vybe-test: js/generators_advanced/generator_symbol_iterator_makes_reusable
// origin: languages/js/tests/js/test_generators_advanced.rs

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

const iterable = {
  [Symbol.iterator]: function* () { yield "x"; yield "y"; yield "z"; }
};
const r1 = [...iterable].join(",");
const r2 = [...iterable].join(",");
__check(__line(r1), "x,y,z");
__check(__line(r1 === r2), "true");
