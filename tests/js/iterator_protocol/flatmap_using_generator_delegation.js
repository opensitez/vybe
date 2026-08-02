// vybe-test: js/iterator_protocol/flatmap_using_generator_delegation
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

function* flatMap(iter, fn) {
    for (const v of iter) yield* fn(v);
}
const result = [...flatMap([1, 2, 3], n => [n, n * 10])];
console.log(result.join(","));
