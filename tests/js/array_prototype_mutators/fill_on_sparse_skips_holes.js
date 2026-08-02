// vybe-test: js/array_prototype_mutators/fill_on_sparse_skips_holes
// origin: languages/js/tests/js/test_array_prototype_mutators.rs

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

const a=[1,,3]; a.fill(0); __check(__line(1 in a), "true"); __check(__line(2 in a), "true");
