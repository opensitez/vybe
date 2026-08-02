// vybe-test: js/graph_tree_algorithms/flatten_tree_to_list
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

function flatten(root, acc=[]) {
    if (!root) return acc;
    acc.push(root.val);
    flatten(root.left, acc);
    flatten(root.right, acc);
    return acc;
}
const t = { val:1, left:{val:2,left:{val:4,left:null,right:null},right:{val:5,left:null,right:null}}, right:{val:3,left:null,right:null} };
__check(__line(flatten(t).join(",")), "1,2,4,5,3");
