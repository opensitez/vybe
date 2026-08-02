// vybe-test: js/namespace_collision_probes/named_fn_expr_name_binds_only_in_own_body
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

const f = function named() { return 1; };
let out;
try { named(); out = "leaked"; } catch (e) { out = "scoped"; }
__check(__line(out), "scoped");
