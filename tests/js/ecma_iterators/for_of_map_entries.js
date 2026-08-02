// vybe-test: js/ecma_iterators/for_of_map_entries
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

const m = new Map();
m.set("a", 1);
m.set("b", 2);
let count = 0;
for (const [k, v] of m) {
    count += v;
}
console.log(count);
