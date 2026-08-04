// vybe-test: js/graph_tree_algorithms/dijkstra_shortest_path
// origin: languages/js/tests/js/test_graph_tree_algorithms.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

// Output is COLLECTED, not paired. The emitter rewrites every `console.log(a)`
// into `__p(__line(a))` and compares the whole buffer once.
//
// Collection is what makes ASYNC assertable at all — 967 of the 1,860 cases the
// per-print emitter refused were `await` / `then` / `Promise`, where the i-th
// log in the SOURCE is not the i-th line of OUTPUT. The buffer records the
// order things actually ran, so no ordering analysis is needed.
let __buf = "";

function __p(s) {
    __buf += s + "\n";
}

function __pr(s) {
    __buf += s;
}

// The check runs from a `setTimeout(…, 0)` — a MACROtask, so it fires only
// after the microtask queue has fully drained. Measured under Vybe: a program
// logging sync, then a `.then`, then past an `await`, then the timeout,
// collects them in exactly that order, while a statement at the end of the
// script sees an empty buffer.
function __checkLater(want) {
    setTimeout(function () {
        __check(__buf, want);
    }, 0);
}

function __check(got, want) {
    // The final log contributes a trailing newline the expected line vector
    // never carried, so both forms are accepted.
    if (got !== want && got !== want + "\n") {
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
__p(__line(d.A));
__p(__line(d.B));
__p(__line(d.C));
__p(__line(d.D));
__checkLater("0\n1\n3\n4");
