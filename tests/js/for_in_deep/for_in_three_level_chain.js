// vybe-test: js/for_in_deep/for_in_three_level_chain
// origin: languages/js/tests/js/test_for_in_deep.rs

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

const a = { from_a: 1 };
const b = Object.create(a);
b.from_b = 2;
const c = Object.create(b);
c.from_c = 3;
const keys = [];
for (const k in c) keys.push(k);
console.log(keys.includes("from_a"));
console.log(keys.includes("from_b"));
console.log(keys.includes("from_c"));
