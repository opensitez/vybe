// vybe-test: js/iterator_protocol/iterable_can_be_iterated_multiple_times
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
        let i = 0;
        return { next() { return i < 3 ? { value: i++, done: false } : { done: true }; } };
    }
};
const r1 = [...iterable];
const r2 = [...iterable];
__check(__line(r1.join(",")), "0,1,2");
__check(__line(r2.join(",")), "0,1,2");
