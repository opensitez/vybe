// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_symbol_iterator_return_method_early_break
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

let cleanedUp = false;
const iterObj = {
    [Symbol.iterator]() {
        return {
            next() { return { value: 1, done: false }; },
            return() {
                cleanedUp = true;
                return { done: true };
            }
        };
    }
};
for (const val of iterObj) {
    if (val === 1) break; // Early break triggers return() method on iterator!
}
console.log(cleanedUp);
