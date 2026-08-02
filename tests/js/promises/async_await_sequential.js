// vybe-test: js/promises/async_await_sequential
// origin: languages/js/tests/js/test_promises.rs

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

async function step(n) {
    return n * 2;
}
async function main() {
    let a = await step(1);
    let b = await step(a);
    let c = await step(b);
    console.log(c);
}
main();
