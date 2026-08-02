// vybe-test: js/array_from_patterns/array_from_custom_iterator
// origin: languages/js/tests/js/test_array_from_patterns.rs

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
        let n = 0;
        return { next() { return n < 3 ? { value: n++, done: false } : { done: true }; } };
    }
};
const arr = Array.from(iterable);
__check(__line(arr.join(",")), "0,1,2");
