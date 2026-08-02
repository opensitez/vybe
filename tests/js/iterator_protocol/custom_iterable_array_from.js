// vybe-test: js/iterator_protocol/custom_iterable_array_from
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

const obj = {
    [Symbol.iterator]() {
        let i = 0;
        return { next() { return i < 4 ? { value: i * i, done: false, } : { done: true }; i++ } };
    }
};
// Simpler approach
function* gen() { for (let i = 0; i < 4; i++) yield i * i; }
const arr = Array.from(gen());
console.log(arr.join(","));
