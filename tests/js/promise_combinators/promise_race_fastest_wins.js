// vybe-test: js/promise_combinators/promise_race_fastest_wins
// origin: languages/js/tests/js/test_promise_combinators.rs

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
    const fast = Promise.resolve("fast");
    const slow = new Promise(resolve => setTimeout(() => resolve("slow"), 100));
    const result = await Promise.race([slow, fast]);
    console.log(result);
}
main();
