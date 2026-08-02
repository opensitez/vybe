// vybe-test: js/scope_closures_edge/for_in_const_binding_per_iteration
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
for (const k in { a: 1, b: 2 }) {
    fns.push(() => k);
}
console.log(fns.map(f => f()).join(","));
