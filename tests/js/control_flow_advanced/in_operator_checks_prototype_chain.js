// vybe-test: js/control_flow_advanced/in_operator_checks_prototype_chain
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

const parent = { foo: 1 };
const child = Object.create(parent);
__check(__line("foo" in child), "true");
__check(__line("bar" in child), "false");
