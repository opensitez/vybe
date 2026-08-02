// vybe-test: js/object_integrity_level_errors/freeze_with_non_configurable_existing
// origin: languages/js/tests/js/test_object_integrity_level_errors.rs

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

const o={}; Object.defineProperty(o,"x",{value:1,configurable:false,writable:true}); Object.freeze(o); __check(__line(Object.isFrozen(o)), "true");
