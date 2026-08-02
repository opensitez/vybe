// vybe-test: js/iterator_protocol/filter_adapter_over_iterable
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

function* filter(iter, pred) {
    for (const v of iter) if (pred(v)) yield v;
}
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const evens = [...filter(range(8), n => n % 2 === 0)];
console.log(evens.join(","));
