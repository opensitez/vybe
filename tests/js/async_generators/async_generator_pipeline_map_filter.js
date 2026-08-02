// vybe-test: js/async_generators/async_generator_pipeline_map_filter
// origin: languages/js/tests/js/test_async_generators.rs

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

async function* range(start, end) {
    for (let i = start; i <= end; i++) yield i;
}
async function* asyncMap(iter, fn) {
    for await (const v of iter) yield fn(v);
}
async function* asyncFilter(iter, pred) {
    for await (const v of iter) if (pred(v)) yield v;
}
async function main() {
    const evensDoubled = asyncFilter(
        asyncMap(range(1, 6), x => x * 2),
        x => x > 4
    );
    const results = [];
    for await (const v of evensDoubled) results.push(v);
    console.log(results.join(","));
}
main();
