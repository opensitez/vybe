// vybe-test: js/scope_closures_edge/var_shared_across_loop_iterations
// origin: languages/js/tests/js/test_scope_closures_edge.rs

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
for (var i = 0; i < 3; i++) {
    fns.push(() => i);
}
// All closures share same var i, which is 3 after loop
console.log(fns.map(f => f()).join(","));
