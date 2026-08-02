// vybe-test: js/symbol_advanced/symbol_to_string_tag_custom
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

class MyCollection {
    get [Symbol.toStringTag]() { return "MyCollection"; }
}
const c = new MyCollection();
__check(__line(Object.prototype.toString.call(c)), "[object MyCollection]");
