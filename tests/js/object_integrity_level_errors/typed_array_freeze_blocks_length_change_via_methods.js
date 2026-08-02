// vybe-test: js/object_integrity_level_errors/typed_array_freeze_blocks_length_change_via_methods
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

try{Object.freeze(new Uint8Array([1]));}catch(e){__check(__line(e instanceof TypeError), "true");}
