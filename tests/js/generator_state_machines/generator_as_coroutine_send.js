// vybe-test: js/generator_state_machines/generator_as_coroutine_send
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

function* counter(start = 0) {
    let n = start;
    while (true) {
        const reset = yield n;
        if (reset === true) n = start;
        else n++;
    }
}
const gen = counter(10);
gen.next(); // start
console.log(gen.next().value);  // 11
console.log(gen.next().value);  // 12
console.log(gen.next(true).value); // reset to 10
console.log(gen.next().value);  // 11
