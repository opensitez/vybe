// vybe-test: js/array_prototype_mutators/copywithin_end_omitted
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

const a=[1,2,3,4,5]; a.copyWithin(0,2); __check(__line(a.join(",")), "3,4,5,4,5");
