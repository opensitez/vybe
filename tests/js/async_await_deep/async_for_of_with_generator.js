// vybe-test: js/async_await_deep/async_for_of_with_generator
// origin: languages/js/tests/js/test_async_await_deep.rs

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

async function* numbers() {
    yield 1;
    yield 2;
    yield 3;
}
async function main() {
    const sum = [];
    for await (const n of numbers()) {
        sum.push(n);
    }
    console.log(sum.join(","));
}
main();
