// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_preserves_iterator_return_method
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

let returned = false;
const customIter = {
    next() { return { value: 1, done: false }; },
    return() { returned = true; return { done: true }; }
};
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(customIter);
    if (typeof wrapped.return === "function") wrapped.return();
    console.log(returned);
} else {
    console.log("true");
}
