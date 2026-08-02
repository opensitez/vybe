// vybe-test: js/graph_tree_algorithms/detect_cycle
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

function hasCycle(graph) {
    const visited = new Set(), stack = new Set();
    function dfs(node) {
        if (stack.has(node)) return true;
        if (visited.has(node)) return false;
        visited.add(node); stack.add(node);
        for (const neighbor of (graph[node] || [])) {
            if (dfs(neighbor)) return true;
        }
        stack.delete(node);
        return false;
    }
    return Object.keys(graph).some(n => dfs(n));
}
console.log(hasCycle({ A: ["B"], B: ["C"], C: ["A"] }));
console.log(hasCycle({ A: ["B"], B: ["C"], C: [] }));
