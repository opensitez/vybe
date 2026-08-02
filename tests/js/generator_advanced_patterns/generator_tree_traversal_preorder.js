// vybe-test: js/generator_advanced_patterns/generator_tree_traversal_preorder
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

function* preorder(node) {
    if (!node) return;
    yield node.val;
    yield* preorder(node.left);
    yield* preorder(node.right);
}
const tree = {
    val: 1,
    left: { val: 2, left: { val: 4, left: null, right: null }, right: null },
    right: { val: 3, left: null, right: null }
};
__check(__line([...preorder(tree)].join(",")), "1,2,4,3");
