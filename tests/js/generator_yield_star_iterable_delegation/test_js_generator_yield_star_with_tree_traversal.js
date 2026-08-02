// vybe-test: js/generator_yield_star_iterable_delegation/test_js_generator_yield_star_with_tree_traversal
// origin: languages/js/tests/js/test_js_generator_yield_star_iterable_delegation.rs

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

const tree = {
    val: 1,
    children: [
        { val: 2, children: [] },
        { val: 3, children: [{ val: 4, children: [] }] }
    ]
};
function* traverse(node) {
    yield node.val;
    for (const child of node.children) {
        yield* traverse(child);
    }
}
console.log([...traverse(tree)].join(","));
