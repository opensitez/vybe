// vybe-test: js/prototype_chain_advanced/null_prototype_has_no_tostring
// origin: languages/js/tests/js/test_prototype_chain_advanced.rs

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

const bare = Object.create(null);
bare.x = 1;
__check(__line(Object.getPrototypeOf(bare) === null), "true");
__check(__line(Object.getPrototypeOf(bare)), "null");
