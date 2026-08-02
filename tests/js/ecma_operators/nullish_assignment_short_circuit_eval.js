// vybe-test: js/ecma_operators/nullish_assignment_short_circuit_eval
// origin: languages/js/tests/js/test_ecma_operators.rs

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

let sideEffect = false;
let val = "initial";
val ??= (sideEffect = true, "fallback");
__check(__line(val), "initial");
__check(__line(sideEffect), "false");
