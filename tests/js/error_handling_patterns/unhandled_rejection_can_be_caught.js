// vybe-test: js/error_handling_patterns/unhandled_rejection_can_be_caught
// origin: languages/js/tests/js/test_error_handling_patterns.rs

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
        Promise.reject(new Error("r1")),
        Promise.resolve("ok"),
        Promise.reject(new Error("r3")),
    ]);
    const errors = results.filter(r => r.status === "rejected").map(r => r.reason.message);
    console.log(errors.join(","));
}
main();
