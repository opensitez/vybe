// vybe-test: js/iterator_protocol/custom_iterable_with_for_of
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
        return {
            next() {
                return i < 3
                    ? { value: i++, done: false }
                    : { value: undefined, done: true };
            }
        };
    }
};
const results = [];
for (const v of iterable) results.push(v);
console.log(results.join(","));
