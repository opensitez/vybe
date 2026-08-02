// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_already_exhausted_iterator
// origin: languages/js/tests/js/test_js_iterator_from_protocol_wrapping.rs

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

const arr = [1];
const iter = arr[Symbol.iterator]();
iter.next(); // Exhaust iter
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(iter);
    console.log([...wrapped].length);
} else {
    console.log("0");
}
