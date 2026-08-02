// vybe-test: js/map_set_deep_patterns/set_as_graph_adjacency
// origin: languages/js/tests/js/test_map_set_deep_patterns.rs

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

class Graph {
    #adj = new Map();
    addEdge(a, b) {
        if (!this.#adj.has(a)) this.#adj.set(a, new Set());
        if (!this.#adj.has(b)) this.#adj.set(b, new Set());
        this.#adj.get(a).add(b);
        this.#adj.get(b).add(a);
    }
    neighbors(node) { return [...(this.#adj.get(node) ?? [])].sort(); }
    hasEdge(a, b) { return (this.#adj.get(a) ?? new Set()).has(b); }
}
const g = new Graph();
g.addEdge("A", "B"); g.addEdge("A", "C"); g.addEdge("B", "C");
__check(__line(g.neighbors("A").join(",")), "B,C");
__check(__line(g.hasEdge("B", "C")), "true");
__check(__line(g.hasEdge("A", "D")), "false");
