// vybe-test: js/generators_advanced/generator_return_statement_value
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

function* gen() { yield 1; yield 2; return "final"; }
const g = gen();
g.next(); g.next();
const { value, done } = g.next();
__check(__line(value, done), "final true");
