// vybe-test: js/generator_advanced_patterns/generator_pipeline_flattening
// origin: languages/js/tests/js/test_generator_advanced_patterns.rs

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

function* flatten(arr, depth = 1) {
    for (const item of arr) {
        if (Array.isArray(item) && depth > 0) yield* flatten(item, depth - 1);
        else yield item;
    }
}
const nested = [1, [2, [3, [4]]], 5];
// JSON.stringify (not join): join would render the still-nested [4] as
// "4" (Array.prototype.join calls toString on elements), hiding the very
// thing this test checks — that depth-limited flatten leaves [4] nested.
console.log(JSON.stringify([...flatten(nested, 2)]));
