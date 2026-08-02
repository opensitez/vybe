// vybe-test: js/scope_closures_edge/let_per_iteration_in_for_loop
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
for (let i = 0; i < 3; i++) {
    fns.push(() => i);
}
console.log(fns.map(f => f()).join(","));
