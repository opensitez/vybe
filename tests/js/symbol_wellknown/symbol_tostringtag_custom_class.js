// vybe-test: js/symbol_wellknown/symbol_tostringtag_custom_class
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

class MyBuffer {
  get [Symbol.toStringTag]() { return "MyBuffer"; }
}
const buf = new MyBuffer();
__check(__line(Object.prototype.toString.call(buf)), "[object MyBuffer]");
