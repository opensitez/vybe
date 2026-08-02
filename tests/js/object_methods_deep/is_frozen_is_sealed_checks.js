// vybe-test: js/object_methods_deep/is_frozen_is_sealed_checks
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const frozen = Object.freeze({});
const sealed = Object.seal({});
const plain = {};
__check(__line(Object.isFrozen(frozen)), "true");
__check(__line(Object.isSealed(sealed)), "true");
__check(__line(Object.isFrozen(plain)), "false");
__check(__line(Object.isSealed(plain)), "false");
