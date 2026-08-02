// vybe-test: js/symbol_wellknown/symbol_iterator_used_in_for_of
// origin: languages/js/tests/js/test_symbol_wellknown.rs

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

const steps = {
  [Symbol.iterator]() {
    let n = 0;
    return { next() { return n < 3 ? { value: n++, done: false } : { done: true }; } };
  }
};
const out = [];
for (const s of steps) out.push(s);
console.log(out.join(","));
