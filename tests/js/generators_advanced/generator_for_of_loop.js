// vybe-test: js/generators_advanced/generator_for_of_loop
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

function* evens(limit) {
  for (let i = 2; i <= limit; i += 2) yield i;
}
const result = [];
for (const v of evens(10)) result.push(v);
console.log(result.join(","));
