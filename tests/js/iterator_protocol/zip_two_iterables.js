// vybe-test: js/iterator_protocol/zip_two_iterables
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

function* zip(a, b) {
    const ia = a[Symbol.iterator]();
    const ib = b[Symbol.iterator]();
    while (true) {
        const ra = ia.next(), rb = ib.next();
        if (ra.done || rb.done) break;
        yield [ra.value, rb.value];
    }
}
const pairs = [...zip([1, 2, 3], ["a", "b", "c"])];
console.log(pairs.map(([n, l]) => n + l).join(","));
