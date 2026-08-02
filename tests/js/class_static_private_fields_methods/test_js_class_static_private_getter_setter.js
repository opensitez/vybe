// vybe-test: js/class_static_private_fields_methods/test_js_class_static_private_getter_setter
// origin: languages/js/tests/js/test_js_class_static_private_fields_methods.rs

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

class State {
    static #count = 0;
    static get #counter() { return State.#count; }
    static set #counter(v) { State.#count = v; }

    static increment() {
        State.#counter = State.#counter + 1;
    }
    static getVal() { return State.#counter; }
}
State.increment();
State.increment();
__check(__line(State.getVal()), "2");
