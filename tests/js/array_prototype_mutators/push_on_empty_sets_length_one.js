// vybe-test: js/array_prototype_mutators/push_on_empty_sets_length_one
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

const a=[]; __check(__line(a.push("x")), "1"); __check(__line(a[0]), "x");
