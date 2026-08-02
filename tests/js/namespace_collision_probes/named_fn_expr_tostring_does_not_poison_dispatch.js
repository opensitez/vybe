// vybe-test: js/namespace_collision_probes/named_fn_expr_tostring_does_not_poison_dispatch
// origin: languages/js/tests/js/test_namespace_collision_probes.rs

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

const f = function toString() { return 1; };
__check(__line(String({ x: 1 })), "[object Object]");
__check(__line(f()), "1");
