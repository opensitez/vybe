// vybe-test: js/iterator_helpers_deep/iterator_zip
// origin: languages/js/tests/js/test_iterator_helpers_deep.rs

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

function* zipIter(a, b) {
    const iterA = a[Symbol.iterator]();
    const iterB = b[Symbol.iterator]();
    while (true) {
        const rA = iterA.next(), rB = iterB.next();
        if (rA.done || rB.done) break;
        yield [rA.value, rB.value];
    }
}
const result = [...zipIter([1, 2, 3], ["a", "b", "c"])];
console.log(result.map(([a, b]) => a + b).join(","));
