// vybe-test: js/edge_cases_final/array_destructure_iterator_once
// origin: languages/js/tests/js/test_edge_cases_final.rs

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

let iterCount = 0;
const iterable = {
    [Symbol.iterator]() {
        iterCount++;
        let i = 0;
        return { next() { return i < 3 ? { value: i++, done: false } : { done: true }; } };
    }
};
const [a, b, c] = iterable;
__check(__line(a), "0");
__check(__line(c), "2");
__check(__line(iterCount), "1");
