// vybe-test: js/async_generators/custom_async_iterable_via_symbol_async_iterator
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

const asyncIterable = {
    [Symbol.asyncIterator]() {
        let i = 0;
        const data = [10, 20, 30];
        return {
            async next() {
                if (i < data.length) return { value: data[i++], done: false };
                return { value: undefined, done: true };
            }
        };
    }
};
async function main() {
    const results = [];
    for await (const v of asyncIterable) results.push(v);
    console.log(results.join(","));
}
main();
