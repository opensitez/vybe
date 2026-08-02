// vybe-test: js/scope_closures_edge/for_of_let_binding_per_iteration
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
for (const v of [10, 20, 30]) {
    fns.push(() => v);
}
console.log(fns.map(f => f()).join(","));
