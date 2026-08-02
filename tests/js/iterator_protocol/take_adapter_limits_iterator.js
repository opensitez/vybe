// vybe-test: js/iterator_protocol/take_adapter_limits_iterator
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

function* take(iter, n) {
    let count = 0;
    for (const v of iter) {
        if (count++ >= n) break;
        yield v;
    }
}
function* naturals() { let n = 1; while (true) yield n++; }
console.log([...take(naturals(), 5)].join(","));
