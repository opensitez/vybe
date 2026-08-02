// vybe-test: js/iterator_protocol/map_adapter_over_custom_iterable
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

function* map(iter, fn) {
    for (const v of iter) yield fn(v);
}
function* range(n) {
    for (let i = 0; i < n; i++) yield i;
}
const result = [...map(range(4), x => x * x)];
console.log(result.join(","));
