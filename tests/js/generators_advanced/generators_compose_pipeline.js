// vybe-test: js/generators_advanced/generators_compose_pipeline
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

function* map(iter, fn) { for (const v of iter) yield fn(v); }
function* filter(iter, pred) { for (const v of iter) if (pred(v)) yield v; }
function* range(n) { for (let i = 1; i <= n; i++) yield i; }

const pipeline = filter(map(range(10), x => x * x), x => x % 2 === 0);
const result = [...pipeline];
console.log(result.join(","));
