// vybe-test: js/graph_tree_algorithms/connected_components
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

function countComponents(n, edges) {
    const parent = Array.from({length: n}, (_, i) => i);
    function find(x) { return parent[x] === x ? x : parent[x] = find(parent[x]); }
    function union(a, b) { parent[find(a)] = find(b); }
    for (const [a, b] of edges) union(a, b);
    return new Set(Array.from({length: n}, (_, i) => find(i))).size;
}
console.log(countComponents(5, [[0,1],[1,2],[3,4]]));
console.log(countComponents(5, []));
console.log(countComponents(4, [[0,1],[1,2],[2,3]]));
