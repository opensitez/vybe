// vybe-test: js/generator_advanced_patterns/generator_as_lazy_map
// origin: languages/js/tests/js/test_generator_advanced_patterns.rs

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

function* map(fn, gen) {
    for (const v of gen) yield fn(v);
}
function* range(n) { for (let i = 1; i <= n; i++) yield i; }
const squares = [...map(x => x * x, range(5))];
console.log(squares.join(","));
