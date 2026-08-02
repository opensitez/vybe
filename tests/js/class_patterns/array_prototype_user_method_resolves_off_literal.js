// vybe-test: js/class_patterns/array_prototype_user_method_resolves_off_literal
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

Array.prototype.second = function () { return this[1]; };
__check(__line([1, 2, 3].second()), "2");
