// vybe-test: js/functional_fp_patterns/reader_monad_simple
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

class Reader {
    constructor(fn) { this.run = fn; }
    map(fn) { return new Reader(env => fn(this.run(env))); }
    flatMap(fn) { return new Reader(env => fn(this.run(env)).run(env)); }
    static of(v) { return new Reader(_ => v); }
    static ask() { return new Reader(env => env); }
}
const greet = Reader.ask()
    .map(env => env.greeting)
    .map(g => g + " World");
__check(__line(greet.run({ greeting: "Hello" })), "Hello World");
__check(__line(greet.run({ greeting: "Hi" })), "Hi World");
