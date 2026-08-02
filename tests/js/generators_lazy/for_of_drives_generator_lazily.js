// vybe-test: js/generators_lazy/for_of_drives_generator_lazily
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

function* count() { yield 10; yield 20; yield 30; }
for (let v of count()) { console.log(v); }
