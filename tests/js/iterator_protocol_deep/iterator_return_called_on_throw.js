// vybe-test: js/iterator_protocol_deep/iterator_return_called_on_throw
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
            next() { return { value: n++, done: false }; },
            return() { returnCalled = true; return { done: true }; }
        };
    }
};
try {
    for (const v of iterable) {
        if (v === 2) throw new Error("stop");
    }
} catch {}
console.log(returnCalled);
