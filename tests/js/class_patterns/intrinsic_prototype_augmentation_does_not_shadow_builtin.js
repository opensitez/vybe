// vybe-test: js/class_patterns/intrinsic_prototype_augmentation_does_not_shadow_builtin
// origin: languages/js/tests/js/test_class_patterns.rs

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

Number.prototype.doubled = function () { return this * 2; };
__check(__line((5).toFixed(2)), "5.00");
