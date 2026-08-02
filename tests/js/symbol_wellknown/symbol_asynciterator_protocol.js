// vybe-test: js/symbol_wellknown/symbol_asynciterator_protocol
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

async function collect() {
  const results = [];
  const asyncIterable = {
    [Symbol.asyncIterator]() {
      let i = 0;
      return {
        async next() {
          if (i < 3) return { value: i++, done: false };
          return { value: undefined, done: true };
        }
      };
    }
  };
  for await (const val of asyncIterable) results.push(val);
  console.log(results.join(","));
}
collect();
