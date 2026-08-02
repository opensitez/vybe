// vybe-test: js/destructuring_patterns/for_of_destructure_object_entries
// origin: languages/js/tests/js/test_destructuring_patterns.rs

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

const scores = { alice: 95, bob: 87, charlie: 92 };
const results = [];
for (const [name, score] of Object.entries(scores)) {
    results.push(`${name}:${score}`);
}
console.log(results.sort().join(","));
