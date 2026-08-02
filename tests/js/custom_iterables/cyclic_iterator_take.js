// vybe-test: js/custom_iterables/cyclic_iterator_take
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

function* cycle(arr) {
    while (true) yield* arr;
}
function take(n, gen) {
    const result = [];
    for (const v of gen) { result.push(v); if (result.length >= n) break; }
    return result;
}
const colors = take(7, cycle(["red", "green", "blue"]));
console.log(colors.join(","));
