// vybe-test: js/symbol_protocols/well_known_symbol_iterator_gen
// origin: languages/js/tests/js/test_symbol_protocols.rs

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

class InfiniteCounter {
    constructor(start = 0) { this.n = start; }
    [Symbol.iterator]() { return this; }
    next() { return { value: this.n++, done: false }; }
}
const counter = new InfiniteCounter(5);
const first5 = [];
for (const n of counter) { first5.push(n); if (first5.length === 5) break; }
console.log(first5.join(","));
