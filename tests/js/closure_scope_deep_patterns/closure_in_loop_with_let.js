// vybe-test: js/closure_scope_deep_patterns/closure_in_loop_with_let
// origin: languages/js/tests/js/test_closure_scope_deep_patterns.rs

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

const fns = [];
for (let i = 0; i < 5; i++) {
    fns.push(() => i);
}
console.log(fns.map(f => f()).join(","));
