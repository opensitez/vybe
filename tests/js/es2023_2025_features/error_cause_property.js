// vybe-test: js/es2023_2025_features/error_cause_property
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

function fetchData() {
    try { throw new Error("network failure"); }
    catch (e) { throw new Error("Failed to fetch", { cause: e }); }
}
try {
    fetchData();
} catch (e) {
    __check(__line(e.message), "Failed to fetch");
    __check(__line(e.cause.message), "network failure");
}
