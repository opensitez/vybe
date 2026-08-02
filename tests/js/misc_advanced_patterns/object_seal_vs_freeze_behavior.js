// vybe-test: js/misc_advanced_patterns/object_seal_vs_freeze_behavior
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

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

const sealed = Object.seal({ x: 1, y: 2 });
sealed.x = 99;
sealed.z = 3;
delete sealed.x;
__check(__line(sealed.x), "99");
__check(__line("z" in sealed), "false");
__check(__line(Object.isSealed(sealed)), "true");

const frozen = Object.freeze({ a: 1 });
frozen.a = 99;
__check(__line(frozen.a), "1");
__check(__line(Object.isFrozen(frozen)), "true");
