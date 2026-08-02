// vybe-test: js/iterator_protocol/iterator_return_called_on_throw
// origin: languages/js/tests/js/test_iterator_protocol.rs

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

const log = [];
const iterable = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() { return { value: i++, done: false }; },
            return() { log.push("cleanup"); return { done: true }; }
        };
    }
};
try {
    for (const v of iterable) {
        if (v === 1) throw new Error("stop");
    }
} catch {}
console.log(log.join(","));
