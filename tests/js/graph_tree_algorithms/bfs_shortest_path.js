// vybe-test: js/graph_tree_algorithms/bfs_shortest_path
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

function bfs(graph, start, end) {
    const queue = [[start, [start]]];
    const visited = new Set([start]);
    while (queue.length) {
        const [node, path] = queue.shift();
        if (node === end) return path;
        for (const neighbor of (graph[node] || [])) {
            if (!visited.has(neighbor)) {
                visited.add(neighbor);
                queue.push([neighbor, [...path, neighbor]]);
            }
        }
    }
    return null;
}
const g = { A: ["B","C"], B: ["D"], C: ["D","E"], D: ["F"], E: ["F"], F: [] };
const path = bfs(g, "A", "F");
console.log(path[0]);
console.log(path[path.length - 1]);
console.log(path.length <= 4);
