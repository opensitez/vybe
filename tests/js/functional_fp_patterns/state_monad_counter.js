// vybe-test: js/functional_fp_patterns/state_monad_counter
// origin: languages/js/tests/js/test_functional_fp_patterns.rs

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
    constructor(fn) { this.run = fn; }
    map(fn) { return new State(s => { const [v, ns] = this.run(s); return [fn(v), ns]; }); }
    flatMap(fn) { return new State(s => { const [v, ns] = this.run(s); return fn(v).run(ns); }); }
    static of(v) { return new State(s => [v, s]); }
    static get() { return new State(s => [s, s]); }
    static put(s) { return new State(_ => [null, s]); }
}
const increment = State.get().flatMap(n => State.put(n + 1).flatMap(() => State.get()));
const [value, finalState] = increment.flatMap(() => increment).run(0);
__check(__line(value), "2");
__check(__line(finalState), "2");
