// vybe-test: js/destructuring_patterns/destructure_generator_output
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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

function* range(n) { for (let i = 0; i < n; i++) yield i; }
const [a, b, c] = range(5);
console.log(a);
console.log(b);
console.log(c);
