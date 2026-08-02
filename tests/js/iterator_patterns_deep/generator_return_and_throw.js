// vybe-test: js/iterator_patterns_deep/generator_return_and_throw
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

function* gen() {
    try {
        yield 1;
        yield 2;
        yield 3;
    } finally {
        yield "cleanup";
    }
}
const g1 = gen();
__check(__line(g1.next().value), "1");
const ret = g1.return("early");
__check(__line(ret.value), "cleanup");
__check(__line(g1.next().done), "true");
