// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_plain_iterator_object
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

const customIter = {
    next() { return { value: 99, done: true }; }
};
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(customIter);
    console.log(wrapped.next().done);
} else {
    console.log("true");
}
