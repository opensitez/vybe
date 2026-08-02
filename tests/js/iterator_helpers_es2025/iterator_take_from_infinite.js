// vybe-test: js/iterator_helpers_es2025/iterator_take_from_infinite
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

function* naturals() { let n = 1; while (true) yield n++; }
const result = Iterator.from(naturals()).take(5).toArray();
console.log(result.join(","));
