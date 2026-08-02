// vybe-test: js/generator_state_machines/generator_state_machine_early_break_cleanup
// origin: languages/js/tests/js/test_generator_state_machines.rs

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

function* stateMachine() {
    try {
        yield "state1";
        yield "state2";
    } finally {
        console.log("cleaned_up");
    }
}
for (const s of stateMachine()) {
    console.log(s);
    break;
}
