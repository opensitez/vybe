// vybe-test: js/generators_lazy/generator_with_arguments_yields_them
// origin: languages/js/tests/js/test_generators_lazy.rs

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

function* range(n) {
    let i = 0;
    while (i < n) { yield i; i = i + 1; }
}
for (let v of range(3)) { console.log(v); }
