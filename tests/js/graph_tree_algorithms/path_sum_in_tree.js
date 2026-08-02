// vybe-test: js/graph_tree_algorithms/path_sum_in_tree
// origin: languages/js/tests/js/test_graph_tree_algorithms.rs

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

function hasPathSum(root, target) {
    if (!root) return false;
    if (!root.left && !root.right) return root.val === target;
    return hasPathSum(root.left, target - root.val) || hasPathSum(root.right, target - root.val);
}
const t = { val: 5, left: { val: 4, left: { val: 11, left: { val: 7, left:null, right:null }, right: { val: 2, left:null, right:null } }, right: null }, right: { val: 8, left:null, right:null } };
__check(__line(hasPathSum(t, 22)), "true");
__check(__line(hasPathSum(t, 10)), "false");
