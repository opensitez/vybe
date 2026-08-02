// vybe-test: js/generators/generator_yield_star
// origin: languages/js/tests/js/test_generators.rs

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

function* inner() {
    yield 2;
    yield 3;
}
function* outer() {
    yield 1;
    yield* inner();
    yield 4;
}
let results = [];
for (let v of outer()) results.push(v);
console.log(results.join(","));
