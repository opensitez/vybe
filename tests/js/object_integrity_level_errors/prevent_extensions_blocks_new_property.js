// vybe-test: js/object_integrity_level_errors/prevent_extensions_blocks_new_property
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

const o={a:1}; Object.preventExtensions(o); o.b=2; __check(__line("b" in o), "false");
