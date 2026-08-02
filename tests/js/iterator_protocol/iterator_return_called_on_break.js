// vybe-test: js/iterator_protocol/iterator_return_called_on_break
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
            return() { log.push("return called"); return { done: true }; }
        };
    }
};
for (const v of iterable) {
    if (v >= 2) break;
}
console.log(log.join(","));
