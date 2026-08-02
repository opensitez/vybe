// vybe-test: js/async_patterns/async_generator_with_await
// origin: languages/js/tests/js/test_async_patterns.rs

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

async function* fetchItems() {
    yield await Promise.resolve("item1");
    yield await Promise.resolve("item2");
    yield await Promise.resolve("item3");
}
async function main() {
    for await (let item of fetchItems()) {
        console.log(item);
    }
}
main();
