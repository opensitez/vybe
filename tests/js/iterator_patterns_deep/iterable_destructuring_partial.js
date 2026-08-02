// vybe-test: js/iterator_patterns_deep/iterable_destructuring_partial
// origin: languages/js/tests/js/test_iterator_patterns_deep.rs

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

function* gen() { yield 10; yield 20; yield 30; yield 40; }
const [a, b, ...rest] = gen();
__check(__line(a), "10");
__check(__line(b), "20");
__check(__line(rest.join(",")), "30,40");
