// vybe-test: js/graph_tree_algorithms/dijkstra_shortest_path
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

function dijkstra(graph, start) {
    const dist = {}, visited = new Set();
    for (const node of Object.keys(graph)) dist[node] = Infinity;
    dist[start] = 0;
    const nodes = Object.keys(graph);
    while (nodes.some(n => !visited.has(n))) {
        const u = nodes.filter(n => !visited.has(n)).sort((a,b)=>dist[a]-dist[b])[0];
        if (dist[u] === Infinity) break;
        visited.add(u);
        for (const [v, w] of Object.entries(graph[u])) {
            if (dist[u] + w < dist[v]) dist[v] = dist[u] + w;
        }
    }
    return dist;
}
const g = { A:{B:1,C:4}, B:{C:2,D:5}, C:{D:1}, D:{} };
const d = dijkstra(g, "A");
console.log(d.A);
console.log(d.B);
console.log(d.C);
console.log(d.D);
