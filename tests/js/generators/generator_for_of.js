// vybe-test: js/generators/generator_for_of
// origin: languages/js/tests/js/test_generators.rs

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

function* range(start, end) {
    for (let i = start; i <= end; i++) {
        yield i;
    }
}
for (let n of range(1, 5)) {
    console.log(n);
}
