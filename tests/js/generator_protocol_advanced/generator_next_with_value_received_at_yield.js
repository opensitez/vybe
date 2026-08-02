// vybe-test: js/generator_protocol_advanced/generator_next_with_value_received_at_yield
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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

function* dialog() {
    const name = yield "What's your name?";
    const age = yield `Hello ${name}, how old are you?`;
    yield `${name} is ${age} years old`;
}
const g = dialog();
__check(__line(g.next().value), "What's your name?");       // "What's your name?"
__check(__line(g.next("Alice").value), "Hello Alice, how old are you?"); // "Hello Alice, how old are you?"
__check(__line(g.next(30).value), "Alice is 30 years old");      // "Alice is 30 years old"
