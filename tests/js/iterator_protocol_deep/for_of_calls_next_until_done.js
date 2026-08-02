// vybe-test: js/iterator_protocol_deep/for_of_calls_next_until_done
// origin: languages/js/tests/js/test_iterator_protocol_deep.rs

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

let nextCalls = 0;
const it = {
    [Symbol.iterator]() { return this; },
    next() {
        nextCalls++;
        return nextCalls <= 3 ? { value: nextCalls, done: false } : { done: true };
    }
};
const vals = [];
for (const v of it) vals.push(v);
console.log(vals.join(","));
console.log(nextCalls); // includes the final done call
