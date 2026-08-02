// vybe-test: js/ecma_operators/optional_chaining_computed_index
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

const list = [{ name: "a" }, { name: "b" }];
__check(__line(list?.[1]?.name), "b");
__check(__line(list?.[5]?.name), "undefined");
