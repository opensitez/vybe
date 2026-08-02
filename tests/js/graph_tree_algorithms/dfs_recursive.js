// vybe-test: js/graph_tree_algorithms/dfs_recursive
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

function dfs(graph, node, visited = new Set()) {
    if (visited.has(node)) return [];
    visited.add(node);
    const result = [node];
    for (const neighbor of (graph[node] || [])) {
        result.push(...dfs(graph, neighbor, visited));
    }
    return result;
}
const g = { 1: [2, 3], 2: [4], 3: [4], 4: [] };
const visited = dfs(g, 1);
console.log(visited.includes(1));
console.log(visited.includes(4));
console.log(visited.length);
