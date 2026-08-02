// vybe-test: js/promise_combinators/promise_race_rejection_wins_if_first
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
    const rejected = Promise.reject(new Error("first"));
    const resolved = Promise.resolve("second");
    try {
        await Promise.race([rejected, resolved]);
    } catch (e) {
        console.log(e.message);
    }
}
main();
