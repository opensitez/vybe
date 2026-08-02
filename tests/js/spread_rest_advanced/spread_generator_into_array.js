// vybe-test: js/spread_rest_advanced/spread_generator_into_array
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

function* range(n) { for (let i = 0; i < n; i++) yield i; }
const arr = [...range(5)];
console.log(arr.join(","));
