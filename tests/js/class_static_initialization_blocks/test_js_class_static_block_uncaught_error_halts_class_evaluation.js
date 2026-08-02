// vybe-test: js/class_static_initialization_blocks/test_js_class_static_block_uncaught_error_halts_class_evaluation
// origin: languages/js/tests/js/test_js_class_static_initialization_blocks.rs

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

try {
    eval(`
        class Unsafe {
            static { throw new Error("ClassEvalError"); }
        }
    `);
} catch (e) {
    __check(__line(e.message), "ClassEvalError");
}
