// vybe-test: js/async_generators/async_generator_lazy_fetch_simulation
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

const pages = ["page1", "page2", "page3"];
async function* fetchPages() {
    for (const page of pages) {
        const data = await Promise.resolve(page.toUpperCase());
        yield data;
    }
}
async function main() {
    const results = [];
    for await (const page of fetchPages()) results.push(page);
    console.log(results.join(","));
}
main();
