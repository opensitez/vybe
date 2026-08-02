// vybe-test: js/generators_advanced/generator_chained_delegation
// origin: languages/js/tests/js/test_generators_advanced.rs

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

function* a() { yield 1; yield 2; }
function* b() { yield* a(); yield 3; }
function* c() { yield* b(); yield 4; }
__check(__line([...c()].join(",")), "1,2,3,4");
