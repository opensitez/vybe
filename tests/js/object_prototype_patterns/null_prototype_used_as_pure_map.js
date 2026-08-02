// vybe-test: js/object_prototype_patterns/null_prototype_used_as_pure_map
// origin: languages/js/tests/js/test_object_prototype_patterns.rs

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

const counts = Object.create(null);
const words = ["apple", "banana", "apple", "cherry", "banana", "apple"];
for (const w of words) counts[w] = (counts[w] ?? 0) + 1;
console.log(counts.apple);
console.log(counts.banana);
console.log(counts.cherry);
