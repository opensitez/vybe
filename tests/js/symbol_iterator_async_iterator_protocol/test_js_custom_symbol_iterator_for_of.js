// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_custom_symbol_iterator_for_of
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

const customCollection = {
    items: [10, 20, 30],
    [Symbol.iterator]() {
        let idx = 0;
        const items = this.items;
        return {
            next() {
                if (idx < items.length) {
                    return { value: items[idx++], done: false };
                }
                return { value: undefined, done: true };
            }
        };
    }
};
const res = [];
for (const val of customCollection) res.push(val);
console.log(res.join(","));
