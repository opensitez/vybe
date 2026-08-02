// vybe-test: js/async_concurrency_patterns/async_compose
// origin: languages/js/tests/js/test_async_concurrency_patterns.rs

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

const asyncPipe = (...fns) => x => fns.reduce(async (p, f) => f(await p), Promise.resolve(x));
async function main() {
    const process = asyncPipe(
        async x => x + 1,
        async x => x * 2,
        async x => x.toString()
    );
    console.log(await process(5));
    console.log(await process(10));
}
main();
