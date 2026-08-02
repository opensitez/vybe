// vybe-test: js/graph_tree_algorithms/level_order_bfs
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

function levelOrder(root) {
    if (!root) return [];
    const result = [], queue = [root];
    while (queue.length) {
        const level = [], n = queue.length;
        for (let i = 0; i < n; i++) {
            const node = queue.shift();
            level.push(node.val);
            if (node.left) queue.push(node.left);
            if (node.right) queue.push(node.right);
        }
        result.push(level);
    }
    return result;
}
const t = { val:1, left:{val:2,left:{val:4,left:null,right:null},right:null}, right:{val:3,left:null,right:null} };
const levels = levelOrder(t);
console.log(levels[0].join(","));
console.log(levels[1].join(","));
console.log(levels[2].join(","));
