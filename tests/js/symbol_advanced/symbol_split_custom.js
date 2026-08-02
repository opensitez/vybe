// vybe-test: js/symbol_advanced/symbol_split_custom
// origin: languages/js/tests/js/test_symbol_advanced.rs

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

class CaseInsensitiveSplit {
    constructor(sep) { this.sep = sep.toLowerCase(); }
    [Symbol.split](str) {
        return str.toLowerCase().split(this.sep);
    }
}
const result = "Hello-WORLD-foo".split(new CaseInsensitiveSplit("-"));
__check(__line(result.join(",")), "hello,world,foo");
