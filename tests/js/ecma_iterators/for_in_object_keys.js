// vybe-test: js/ecma_iterators/for_in_object_keys
// origin: languages/js/tests/js/test_ecma_iterators.rs

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

const obj = { x: 1, y: 2, z: 3 };
const keys = [];
for (const k in obj) {
    keys.push(k);
}
console.log(keys.length);
