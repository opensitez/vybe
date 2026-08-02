// vybe-test: js/control_flow_advanced/if_statement_block_scope_isolated_between_branches
// origin: languages/js/tests/js/test_control_flow_advanced.rs

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

const out = [];
if (false) {
    const branch = "if";
    out.push(branch);
} else {
    const branch = "else";
    out.push(branch);
}
let leaked = false;
try {
    branch;
} catch (e) {
    leaked = e instanceof ReferenceError;
}
out.push(String(leaked));
__check(__line(out.join("|")), "else|true");
