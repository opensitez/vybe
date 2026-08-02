// vybe-test: js/generator_protocol_advanced/generator_lazy_map_filter
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

function* lazyMap(iter, fn) {
    for (const v of iter) yield fn(v);
}
function* lazyFilter(iter, pred) {
    for (const v of iter) if (pred(v)) yield v;
}
function* range(n) { for (let i = 0; i < n; i++) yield i; }
function take(iter, n) {
    const result = [];
    for (const v of iter) { result.push(v); if (result.length >= n) break; }
    return result;
}

const pipeline = lazyFilter(
    lazyMap(range(100), x => x * x),
    x => x % 2 === 0
);
console.log(take(pipeline, 5).join(","));
