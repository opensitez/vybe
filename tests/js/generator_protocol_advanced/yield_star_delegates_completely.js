// vybe-test: js/generator_protocol_advanced/yield_star_delegates_completely
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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

function* inner() { yield "a"; yield "b"; }
function* outer() { yield* inner(); yield "c"; }
__check(__line([...outer()].join(",")), "a,b,c");
