// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_symbol_iterator_array_spread
// origin: languages/js/tests/js/test_js_symbol_iterator_async_iterator_protocol.rs

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

const range = {
    from: 1, to: 3,
    *[Symbol.iterator]() {
        for (let i = this.from; i <= this.to; i++) yield i;
    }
};
console.log([...range].join("-"));
