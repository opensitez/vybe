// vybe-test: js/generator_protocol_advanced/recursive_tree_traversal_via_generator
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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

function* walk(node) {
    yield node.value;
    if (node.left) yield* walk(node.left);
    if (node.right) yield* walk(node.right);
}
const tree = {
    value: 1,
    left: { value: 2, left: { value: 4, left: null, right: null }, right: null },
    right: { value: 3, left: null, right: null }
};
__check(__line([...walk(tree)].join(",")), "1,2,4,3");
