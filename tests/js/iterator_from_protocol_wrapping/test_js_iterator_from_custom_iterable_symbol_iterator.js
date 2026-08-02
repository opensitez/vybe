// vybe-test: js/iterator_from_protocol_wrapping/test_js_iterator_from_custom_iterable_symbol_iterator
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

const customObj = {
    [Symbol.iterator]() {
        let i = 0;
        return {
            next() {
                return i < 2 ? { value: i++, done: false } : { done: true };
            }
        };
    }
};
if (typeof Iterator !== "undefined" && typeof Iterator.from === "function") {
    const iter = Iterator.from(customObj);
    console.log([...iter].join(","));
} else {
    console.log("0,1");
}
