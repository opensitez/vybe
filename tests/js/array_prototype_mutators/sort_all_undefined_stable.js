// vybe-test: js/array_prototype_mutators/sort_all_undefined_stable
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

const a=[undefined,1,undefined]; a.sort(); __check(__line(a[0]===undefined), "false"); __check(__line(a[2]), "undefined");
