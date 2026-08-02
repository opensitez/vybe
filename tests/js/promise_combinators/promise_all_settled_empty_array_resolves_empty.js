// vybe-test: js/promise_combinators/promise_all_settled_empty_array_resolves_empty
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
    const res = await Promise.allSettled([]);
    console.log(Array.isArray(res) + "|" + res.length);
}
main();
