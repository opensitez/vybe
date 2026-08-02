// vybe-test: js/generators_advanced/generator_take_helper
// origin: languages/js/tests/js/test_generators_advanced.rs

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

function* naturals() { let n = 1; while (true) yield n++; }
function take(gen, n) {
  const result = [];
  for (const v of gen) { result.push(v); if (result.length === n) break; }
  return result;
}
console.log(take(naturals(), 5).join(","));
