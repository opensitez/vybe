// vybe-test: js/optional_chaining_edge/optional_chain_short_circuits_member_access
// origin: languages/js/tests/js/test_optional_chaining_edge.rs

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

let count = 0;
function sideEffect() { count++; return 1; }
const obj = null;
obj?.prop[sideEffect()]; // sideEffect should NOT be called
__check(__line(count), "0");
