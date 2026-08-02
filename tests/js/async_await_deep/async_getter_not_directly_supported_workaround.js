// vybe-test: js/async_await_deep/async_getter_not_directly_supported_workaround
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

class Loader {
    async load() {
        return await Promise.resolve("data");
    }
}
async function main() {
    const loader = new Loader();
    console.log(await loader.load());
}
main();
