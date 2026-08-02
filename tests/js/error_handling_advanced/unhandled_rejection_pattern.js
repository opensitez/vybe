// vybe-test: js/error_handling_advanced/unhandled_rejection_pattern
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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
        Promise.reject(new Error("fail")),
        Promise.resolve(3),
    ]);
    const statuses = results.map(r => r.status);
    console.log(statuses.join(","));
    console.log(results[1].reason.message);
}
main();
