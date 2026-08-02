// vybe-test: js/iterator_patterns_deep/for_of_with_index_via_entries
// origin: languages/js/tests/js/test_iterator_patterns_deep.rs

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

const items = ["a", "b", "c", "d"];
const indexed = [];
for (const [i, v] of items.entries()) indexed.push(`${i}:${v}`);
console.log(indexed.join(","));
