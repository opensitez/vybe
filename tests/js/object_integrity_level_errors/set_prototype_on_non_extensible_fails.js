// vybe-test: js/object_integrity_level_errors/set_prototype_on_non_extensible_fails
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

const o=Object.preventExtensions({}); try{Object.setPrototypeOf(o,{});}catch(e){__check(__line(e instanceof TypeError), "true");}
