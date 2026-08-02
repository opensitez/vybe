// vybe-test: js/graph_tree_algorithms/topological_sort
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

function topoSort(nodes, edges) {
    const inDegree = {};
    const adj = {};
    for (const n of nodes) { inDegree[n] = 0; adj[n] = []; }
    for (const [from, to] of edges) { adj[from].push(to); inDegree[to]++; }
    const queue = nodes.filter(n => inDegree[n] === 0);
    const result = [];
    while (queue.length) {
        const node = queue.shift();
        result.push(node);
        for (const next of adj[node]) {
            if (--inDegree[next] === 0) queue.push(next);
        }
    }
    return result;
}
const order = topoSort([1,2,3,4,5], [[1,3],[2,3],[3,4],[3,5]]);
console.log(order.indexOf(1) < order.indexOf(3));
console.log(order.indexOf(3) < order.indexOf(4));
console.log(order.length);
