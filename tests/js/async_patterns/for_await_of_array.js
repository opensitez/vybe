// vybe-test: js/async_patterns/for_await_of_array
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

async function main() {
    let promises = [
        Promise.resolve("a"),
        Promise.resolve("b"),
        Promise.resolve("c")
    ];
    for await (let val of promises) {
        console.log(val);
    }
}
main();
