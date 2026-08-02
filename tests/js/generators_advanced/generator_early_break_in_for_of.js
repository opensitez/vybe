// vybe-test: js/generators_advanced/generator_early_break_in_for_of
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

const visited = [];
function* gen() {
  try {
    yield 1; yield 2; yield 3;
  } finally {
    visited.push("cleanup");
  }
}
for (const v of gen()) {
  visited.push(v);
  if (v === 2) break;
}
console.log(visited.join(","));
