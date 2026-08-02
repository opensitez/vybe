// vybe-test: js/iterator_protocol_deep/destructuring_uses_iterator
// origin: languages/js/tests/js/test_iterator_protocol_deep.rs

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

function* gen() { yield 1; yield 2; yield 3; }
const [a, b, c] = gen();
__check(__line(a), "1");
__check(__line(b), "2");
__check(__line(c), "3");
