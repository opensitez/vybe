// vybe-test: js/generators_advanced/yield_star_string_iterates_chars
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

function* charGen() { yield* "abc"; }
__check(__line([...charGen()].join(",")), "a,b,c");
