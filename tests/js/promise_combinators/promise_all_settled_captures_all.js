// vybe-test: js/promise_combinators/promise_all_settled_captures_all
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
    const results = await Promise.allSettled([
        Promise.resolve(1),
        Promise.reject("err"),
        Promise.resolve(3),
    ]);
    console.log(results[0].status);
    console.log(results[1].status);
    console.log(results[1].reason);
    console.log(results[2].value);
}
main();
