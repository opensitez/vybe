// vybe-test: js/object_integrity_level_errors/seal_blocks_add_and_delete
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

const o={a:1}; Object.seal(o); o.b=2; delete o.a; __check(__line("a" in o), "true");__check(__line("b" in o), "false");
