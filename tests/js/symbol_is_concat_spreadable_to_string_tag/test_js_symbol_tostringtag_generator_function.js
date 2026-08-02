// vybe-test: js/symbol_is_concat_spreadable_to_string_tag/test_js_symbol_tostringtag_generator_function
// origin: languages/js/tests/js/test_js_symbol_is_concat_spreadable_to_string_tag.rs

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

function* myGenerator() {}
__check(__line(Object.prototype.toString.call(myGenerator())), "[object Generator]");
