// vybe-test: js/iterator_helpers_es2025/iterator_flatmap_with_generator
// origin: languages/js/tests/js/test_iterator_helpers_es2025.rs

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

const result = Iterator.from([2, 3]).flatMap(function*(x) {
    for (let i = 1; i <= x; i++) yield i;
}).toArray();
console.log(result.join(","));
