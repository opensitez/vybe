// vybe-test: js/async_error_handling/async_pipeline_pattern
// origin: languages/js/tests/js/test_async_error_handling.rs

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

const steps = [
    async (x) => x + 1,
    async (x) => x * 2,
    async (x) => x - 3,
];

async function pipeline(input, fns) {
    let value = input;
    for (const fn of fns) value = await fn(value);
    return value;
}

pipeline(5, steps).then(v => console.log(v)); // (5+1)*2-3 = 9
