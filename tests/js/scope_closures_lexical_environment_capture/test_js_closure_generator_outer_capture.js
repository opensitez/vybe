// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_generator_outer_capture
// origin: languages/js/tests/js/test_js_scope_closures_lexical_environment_capture.rs

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

function makeSeqGenerator(step) {
    return function*() {
        let current = 0;
        while (true) {
            yield (current += step);
        }
    };
}
const gen = makeSeqGenerator(5)();
console.log(`${gen.next().value}:${gen.next().value}:${gen.next().value}`);
