// vybe-test: js/custom_iterables/cartesian_product_generator
// origin: languages/js/tests/js/test_custom_iterables.rs

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

function* cartesian(a, b) {
    for (const x of a) for (const y of b) yield [x, y];
}
const pairs = [...cartesian([1, 2], ["a", "b"])];
console.log(pairs.length);
console.log(pairs.map(([x,y]) => x+y).join(","));
