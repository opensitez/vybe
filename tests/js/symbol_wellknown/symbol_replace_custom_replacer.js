// vybe-test: js/symbol_wellknown/symbol_replace_custom_replacer
// origin: languages/js/tests/js/test_symbol_wellknown.rs

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

const replacer = {
  [Symbol.replace](str, replacement) {
    return str.split("x").join(replacement);
  }
};
__check(__line("axbxc".replace(replacer, "-")), "a-b-c");
