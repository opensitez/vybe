// vybe-test: js/scope_closure_patterns/generator_closure_state
// origin: languages/js/tests/js/test_scope_closure_patterns.rs

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

function* statefulGen(start) {
    let n = start;
    while (true) {
        const reset = yield n;
        if (reset) n = start;
        else n++;
    }
}
const gen = statefulGen(0);
console.log(gen.next().value);
console.log(gen.next().value);
console.log(gen.next().value);
console.log(gen.next(true).value); // reset
console.log(gen.next().value);
