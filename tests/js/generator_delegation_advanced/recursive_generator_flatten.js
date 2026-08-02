// vybe-test: js/generator_delegation_advanced/recursive_generator_flatten
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

function* flatten(arr) {
    for (const item of arr) {
        if (Array.isArray(item)) yield* flatten(item);
        else yield item;
    }
}
const nested = [1, [2, [3, [4, [5]]]], 6];
console.log([...flatten(nested)].join(","));
