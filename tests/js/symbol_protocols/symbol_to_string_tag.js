// vybe-test: js/symbol_protocols/symbol_to_string_tag
// origin: languages/js/tests/js/test_symbol_protocols.rs

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
const mc = new MyCollection();
__check(__line(Object.prototype.toString.call(mc)), "[object MyCollection]");
__check(__line(mc.toString()), "[object MyCollection]");
