// vybe-test: js/async_generator_deep/async_generator_yield_star_captures_return_value
// origin: languages/js/tests/js/test_async_generator_deep.rs

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

async function* sub() {
    yield 1;
    return "ret";
}
async function* mainGen() {
    const r = yield* sub();
    yield r;
}
async function main() {
    const res = [];
    for await (const v of mainGen()) res.push(v);
    console.log(res.join(","));
}
main();
