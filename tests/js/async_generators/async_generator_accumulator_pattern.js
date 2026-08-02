// vybe-test: js/async_generators/async_generator_accumulator_pattern
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

async function* accumulate() {
    let sum = 0;
    while (true) {
        const val = yield sum;
        if (val === null) return;
        sum += val;
    }
}
async function main() {
    const g = accumulate();
    await g.next();
    await g.next(10);
    await g.next(20);
    const r = await g.next(30);
    console.log(r.value);
}
main();
