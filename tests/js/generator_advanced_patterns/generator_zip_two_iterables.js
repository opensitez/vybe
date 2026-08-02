// vybe-test: js/generator_advanced_patterns/generator_zip_two_iterables
// origin: languages/js/tests/js/test_generator_advanced_patterns.rs

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
    const itA = a[Symbol.iterator]();
    const itB = b[Symbol.iterator]();
    while (true) {
        const rA = itA.next();
        const rB = itB.next();
        if (rA.done || rB.done) break;
        yield [rA.value, rB.value];
    }
}
const zipped = [...zip([1, 2, 3], ["a", "b", "c"])];
console.log(zipped.map(([a, b]) => a + b).join(","));
