// vybe-test: js/iterator_protocol/custom_iterable_spread
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

const iterable = {
    [Symbol.iterator]() {
        const vals = [10, 20, 30];
        let i = 0;
        return { next() { return i < vals.length ? { value: vals[i++], done: false } : { done: true }; } };
    }
};
const arr = [...iterable];
__check(__line(arr.join(",")), "10,20,30");
