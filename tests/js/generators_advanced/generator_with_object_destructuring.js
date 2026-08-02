// vybe-test: js/generators_advanced/generator_with_object_destructuring
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

function* pairs() {
  yield { key: "a", val: 1 };
  yield { key: "b", val: 2 };
}
const result = [];
for (const { key, val } of pairs()) result.push(key + "=" + val);
console.log(result.join(","));
