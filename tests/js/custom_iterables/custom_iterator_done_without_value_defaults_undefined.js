// vybe-test: js/custom_iterables/custom_iterator_done_without_value_defaults_undefined
// origin: languages/js/tests/js/test_custom_iterables.rs

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
        let count = 0;
        return {
            next() {
                return count++ === 0 ? { value: "a", done: false } : { done: true };
            }
        };
    }
};
const [a, b] = iterable;
__check(__line(a + "|" + b), "a|undefined");
