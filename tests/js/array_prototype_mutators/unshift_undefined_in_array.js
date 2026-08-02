// vybe-test: js/array_prototype_mutators/unshift_undefined_in_array
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

const a=[1,2]; a.unshift(undefined); __check(__line(a.length), "3"); __check(__line(0 in a), "true");
