// vybe-test: js/control_flow_advanced/catch_optional_binding_omitted
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

let caught = false;
try { throw new Error("x"); } catch { caught = true; }
__check(__line(caught), "true");
