// vybe-test: js/object_integrity_level_errors/strict_write_on_frozen_throws_in_strict
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

"use strict"; const o=Object.freeze({x:1}); try{o.x=2;}catch(e){__check(__line(e instanceof TypeError), "true");}
