// vybe-test: js/scope_closures_edge/switch_shares_let_binding_across_cases
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

switch (1) {
    case 1:
        let v = "from 1";
        // fall through
    case 2:
        // v visible here (same block)
        __check(__line(v), "from 1");
        break;
}
