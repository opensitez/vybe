// vybe-test: js/symbol_wellknown/symbol_iterator_custom_iterable
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

const range = {
  from: 1, to: 5,
  [Symbol.iterator]() {
    let cur = this.from;
    const last = this.to;
    return {
      next() {
        return cur <= last ? { value: cur++, done: false } : { value: undefined, done: true };
      }
    };
  }
};
const result = [...range];
__check(__line(result.join(",")), "1,2,3,4,5");
