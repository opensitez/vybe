// vybe-test: js/iterator_protocol/custom_iterable_destructuring
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

const range = {
    [Symbol.iterator]() {
        let n = 1;
        return { next() { return n <= 5 ? { value: n++, done: false } : { done: true }; } };
    }
};
const [a, b, c] = range;
__check(__line(a), "1");
__check(__line(b), "2");
__check(__line(c), "3");
