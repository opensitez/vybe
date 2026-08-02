// vybe-test: js/generator_delegation_advanced/generator_pipeline
// origin: languages/js/tests/js/test_generator_delegation_advanced.rs

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

function* naturals(n = Infinity) {
    let i = 0;
    while (i < n) yield i++;
}
function* map(gen, fn) { for (const v of gen) yield fn(v); }
function* filter(gen, pred) { for (const v of gen) if (pred(v)) yield v; }
function* take(gen, n) { let i = 0; for (const v of gen) { if (i++ >= n) break; yield v; } }

const result = [...take(filter(map(naturals(), x => x*x), x => x % 2 === 0), 5)];
console.log(result.join(","));
