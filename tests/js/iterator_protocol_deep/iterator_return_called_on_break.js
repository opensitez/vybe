// vybe-test: js/iterator_protocol_deep/iterator_return_called_on_break
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

let returnCalled = false;
const iterable = {
    [Symbol.iterator]() {
        let n = 0;
        return {
            next() { return n < 10 ? { value: n++, done: false } : { done: true }; },
            return() { returnCalled = true; return { done: true }; }
        };
    }
};
for (const v of iterable) {
    if (v === 2) break;
}
console.log(returnCalled);
