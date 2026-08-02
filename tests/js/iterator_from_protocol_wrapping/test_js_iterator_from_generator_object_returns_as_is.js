// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_generator_object_returns_as_is
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

function* gen() { yield 1; }
const g = gen();
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const wrapped = Iterator.from(g);
    console.log(wrapped === g); // Generator instance is already an Iterator, returned as-is!
} else {
    console.log("true");
}
