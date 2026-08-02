// vybe-test: js/module_import_patterns/async_function_simulates_top_level_await
// origin: languages/js/tests/js/test_module_import_patterns.rs

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

// In module scripts, top-level await is valid
// In non-module context, we simulate with async IIFE
(async () => {
    const data = await Promise.resolve("fetched");
    console.log(data);
})();
