// vybe-test: js/generator_state_machines/depth_first_search_generator
// origin: languages/js/tests/js/test_generator_state_machines.rs

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

function* dfs(graph, start, visited = new Set()) {
    if (visited.has(start)) return;
    visited.add(start);
    yield start;
    for (const neighbor of (graph[start] || [])) {
        yield* dfs(graph, neighbor, visited);
    }
}
const graph = {
    A: ["B", "C"],
    B: ["D"],
    C: ["D", "E"],
    D: [],
    E: []
};
const order = [...dfs(graph, "A")];
console.log(order.join(","));
