// vybe-test: js/generator_state_machines/generator_as_observable
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

function* events(data) {
    for (const item of data) {
        if (item > 0) yield { type: "positive", value: item };
        else if (item < 0) yield { type: "negative", value: item };
        else yield { type: "zero", value: 0 };
    }
}
const evts = [...events([1, -2, 0, 3])];
console.log(evts.map(e => e.type).join(","));
console.log(evts.map(e => e.value).join(","));
