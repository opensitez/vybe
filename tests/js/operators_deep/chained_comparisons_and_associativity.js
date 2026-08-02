// vybe-test: js/operators_deep/chained_comparisons_and_associativity
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

__check(__line(1 + 2 * 3 - 4 / 2), "5");
__check(__line(1 < 2 < 3), "true"); // left-to-right: (1 < 2) -> true -> 1 -> 1 < 3
__check(__line(3 < 2 < 1), "true"); // left-to-right: (3 < 2) -> false -> 0 -> 0 < 1
