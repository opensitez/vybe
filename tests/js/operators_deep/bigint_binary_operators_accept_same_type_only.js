// vybe-test: js/operators_deep/bigint_binary_operators_accept_same_type_only
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line((1n + 2n).toString()), "3");
__check(__line((1n << 2n).toString()), "4");
let addMixed = false;
let shiftMixed = false;
try {
    const _ = 1n + 2;
} catch (e) {
    addMixed = true;
}
try {
    const _ = 1n << 2;
} catch (e) {
    shiftMixed = true;
}
__check(__line(`${addMixed}:${shiftMixed}`), "true:true");
