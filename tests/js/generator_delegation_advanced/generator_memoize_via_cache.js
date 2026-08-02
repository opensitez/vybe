// vybe-test: js/generator_delegation_advanced/generator_memoize_via_cache
// origin: languages/js/tests/js/test_generator_delegation_advanced.rs

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

function* uniqueValues(gen) {
    const seen = new Set();
    for (const v of gen) {
        if (!seen.has(v)) { seen.add(v); yield v; }
    }
}
const input = [1, 2, 2, 3, 1, 4, 3, 5];
console.log([...uniqueValues(input)].join(","));
