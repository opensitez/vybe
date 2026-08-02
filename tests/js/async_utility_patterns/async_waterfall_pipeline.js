// vybe-test: js/async_utility_patterns/async_waterfall_pipeline
// origin: languages/js/tests/js/test_async_utility_patterns.rs

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

async function waterfall(fns, initial) {
    return fns.reduce(async (promise, fn) => fn(await promise), Promise.resolve(initial));
}
async function main() {
    const result = await waterfall([
        async x => x + 1,
        async x => x * 2,
        async x => "result: " + x,
    ], 5);
    console.log(result);
}
main();
