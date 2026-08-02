// vybe-test: js/ecma_iterators/continue_skips_current_iteration_only
// origin: languages/js/tests/js/test_ecma_iterators.rs

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

let seen = [];
for (let i = 0; i < 4; i++) {
    if (i === 2) continue;
    seen.push(i);
}
console.log(seen.join(","));
