// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_exhausted_subsequent_next_calls
// origin: languages/js/tests/js/test_js_generator_return_throw_next_state_machine.rs

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

function* gen() { yield 1; }
const g = gen();
g.next(); // value: 1, done: false
g.next(); // value: undefined, done: true
const s3 = g.next(); // value: undefined, done: true
__check(__line(`${s3.value}:${s3.done}`), "undefined:true");
