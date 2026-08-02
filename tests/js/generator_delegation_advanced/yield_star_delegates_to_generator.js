// vybe-test: js/generator_delegation_advanced/yield_star_delegates_to_generator
// origin: languages/js/tests/js/test_generator_delegation_advanced.rs

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

function* inner() { yield 1; yield 2; yield 3; }
function* outer() { yield 0; yield* inner(); yield 4; }
__check(__line([...outer()].join(",")), "0,1,2,3,4");
