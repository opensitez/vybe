// vybe-test: js/array_prototype_mutators/sort_holes_and_undefined_moved_to_end
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

const a = [3, , undefined, 1]; a.sort(); __check(__line(a[0] + "," + a[1] + "|" + (2 in a) + "|" + (3 in a)), "1,3|true|false");
