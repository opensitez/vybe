// vybe-test: js/generator_delegation_advanced/yield_star_three_level_delegation_return_cascade
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

function* g1() { yield 1; return "r1"; }
function* g2() { const r = yield* g1(); yield r; return "r2"; }
function* g3() { const r = yield* g2(); yield r; }
__check(__line([...g3()].join(",")), "1,r1,r2");
