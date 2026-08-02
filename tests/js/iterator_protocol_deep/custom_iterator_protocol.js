// vybe-test: js/iterator_protocol_deep/custom_iterator_protocol
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

function makeCounter(max) {
    let n = 0;
    return {
        [Symbol.iterator]() { return this; },
        next() {
            return n < max ? { value: n++, done: false } : { done: true, value: undefined };
        }
    };
}
__check(__line([...makeCounter(3)].join(",")), "0,1,2");
