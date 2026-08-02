// vybe-test: js/array_prototype_mutators/push_returns_new_length
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

const a=[1,2]; __check(__line(a.push(3,4)), "4"); __check(__line(a.join(",")), "1,2,3,4");
