// vybe-test: js/operators_deep/nullish_coalescing_only_null_undefined
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line(0 ?? "default"), "0");      // 0 — not null/undefined
__check(__line("" ?? "default"), "");     // "" — not null/undefined
__check(__line(false ?? "default"), "false");  // false — not null/undefined
__check(__line(null ?? "default"), "default");   // "default"
__check(__line(undefined ?? "default"), "default"); // "default"
