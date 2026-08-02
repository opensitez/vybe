// vybe-test: js/generator_protocol_advanced/generator_is_both_iterable_and_iterator
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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

function* gen() { yield 1; yield 2; }
const g = gen();
// Generator has Symbol.iterator that returns itself
console.log(g[Symbol.iterator]() === g);
// So it can be used in for...of after partially consuming
g.next(); // consume 1
const remaining = [...g]; // consume rest
console.log(remaining.join(","));
