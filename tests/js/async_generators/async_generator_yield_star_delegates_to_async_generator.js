// vybe-test: js/async_generators/async_generator_yield_star_delegates_to_async_generator
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

async function* inner() {
    yield "x";
    yield "y";
}
async function* outer() {
    yield "start";
    yield* inner();
    yield "end";
}
async function main() {
    const results = [];
    for await (const v of outer()) results.push(v);
    console.log(results.join(","));
}
main();
