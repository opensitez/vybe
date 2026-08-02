// vybe-test: js/promise_combinators/promise_all_rejects_on_first_failure
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
    try {
        await Promise.all([
            Promise.resolve(1),
            Promise.reject("boom"),
            Promise.resolve(3),
        ]);
    } catch (e) {
        console.log(e);
    }
}
main();
