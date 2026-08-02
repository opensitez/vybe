// vybe-test: js/reactive_patterns/stream_accumulator
// origin: languages/js/tests/js/test_reactive_patterns.rs

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

class Stream {
    #buffer = [];
    #subscribers = [];
    push(...values) {
        this.#buffer.push(...values);
        values.forEach(v => this.#subscribers.forEach(fn => fn(v)));
        return this;
    }
    subscribe(fn) { this.#subscribers.push(fn); return this; }
    collect() { return [...this.#buffer]; }
    map(fn) {
        const out = new Stream();
        this.subscribe(v => out.push(fn(v)));
        return out;
    }
    filter(fn) {
        const out = new Stream();
        this.subscribe(v => { if (fn(v)) out.push(v); });
        return out;
    }
}
const s = new Stream();
const evens = s.filter(x => x % 2 === 0).map(x => x * 10);
const results = [];
evens.subscribe(v => results.push(v));
s.push(1, 2, 3, 4, 5);
console.log(results.join(","));
