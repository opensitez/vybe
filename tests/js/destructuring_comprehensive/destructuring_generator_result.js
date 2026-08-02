// vybe-test: js/destructuring_comprehensive/destructuring_generator_result
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

function* entries(obj) {
    for (const [k, v] of Object.entries(obj)) yield [k, v];
}
const obj = { x: 10, y: 20, z: 30 };
const results = [];
for (const [key, val] of entries(obj)) results.push(key + ":" + val);
console.log(results.join(","));
