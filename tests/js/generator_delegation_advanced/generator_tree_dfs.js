// vybe-test: js/generator_delegation_advanced/generator_tree_dfs
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

function* dfs(node) {
    yield node.value;
    if (node.left) yield* dfs(node.left);
    if (node.right) yield* dfs(node.right);
}
const tree = {
    value: 1,
    left: { value: 2, left: { value: 4, left: null, right: null }, right: null },
    right: { value: 3, left: null, right: { value: 5, left: null, right: null } }
};
__check(__line([...dfs(tree)].join(",")), "1,2,4,3,5");
