// vybe-test: js/es2023_2025_features/promise_any_first_resolve
// origin: languages/js/tests/js/test_es2023_2025_features.rs

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

async function test() {
    const p = await Promise.any([
        Promise.reject(1),
        Promise.resolve(2),
        Promise.resolve(3)
    ]);
    console.log(p);
}
test();
