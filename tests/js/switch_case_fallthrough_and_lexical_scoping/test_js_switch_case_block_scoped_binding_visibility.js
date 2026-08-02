// vybe-test: js/switch_case_fallthrough_and_lexical_scoping/test_js_switch_case_block_scoped_binding_visibility
// origin: languages/js/tests/js/test_js_switch_case_fallthrough_and_lexical_scoping.rs

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

const events = [];
switch (1) {
    case 1: {
        const local = "inside";
        events.push(local);
        break;
    }
    default:
        events.push("default");
}

let leaked = false;
try {
    local;
} catch (e) {
    leaked = e instanceof ReferenceError;
}

events.push(String(leaked));
__check(__line(events.join("|")), "inside|true");
