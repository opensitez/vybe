// vybe-test: js/symbol_wellknown/symbol_split_custom_splitter
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

const splitter = {
  [Symbol.split](str) {
    return str.split("").filter(c => c !== " ");
  }
};
const result = "a b c".split(splitter);
__check(__line(result.join(",")), "a,b,c");
