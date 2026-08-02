// vybe-test: js/iterator_protocol_deep/iterable_object_can_be_spread
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

const range = {
    from: 1, to: 5,
    [Symbol.iterator]() {
        let cur = this.from, end = this.to;
        return { next() { return cur <= end ? { value: cur++, done: false } : { done: true }; } };
    }
};
__check(__line([...range].join(",")), "1,2,3,4,5");
