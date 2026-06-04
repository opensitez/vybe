/// Graph and tree algorithms in pure JavaScript
use super::helpers::run_js;

#[test]
fn bfs_shortest_path() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["A", "F", "true"]
    );
}

#[test]
fn dfs_recursive() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["true", "true", "4"]
    );
}

#[test]
fn topological_sort() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["true", "true", "5"]
    );
}

#[test]
fn connected_components() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["2", "5", "1"]
    );
}

#[test]
fn binary_tree_traversal() {
    assert_eq!(
        run_js(
            r#"
class TreeNode {
    constructor(val, left=null, right=null) { this.val=val; this.left=left; this.right=right; }
}
function inorder(root, result=[]) {
    if (!root) return result;
    inorder(root.left, result);
    result.push(root.val);
    inorder(root.right, result);
    return result;
}
const root = new TreeNode(4, new TreeNode(2, new TreeNode(1), new TreeNode(3)), new TreeNode(5));
console.log(inorder(root).join(","));
"#
        ),
        vec!["1,2,3,4,5"]
    );
}

#[test]
fn binary_tree_max_depth() {
    assert_eq!(
        run_js(
            r#"
function maxDepth(root) {
    if (!root) return 0;
    return 1 + Math.max(maxDepth(root.left), maxDepth(root.right));
}
const t = { val: 1, left: { val: 2, left: { val: 4, left: null, right: null }, right: null }, right: { val: 3, left: null, right: null } };
console.log(maxDepth(t));
console.log(maxDepth(null));
"#
        ),
        vec!["3", "0"]
    );
}

#[test]
fn path_sum_in_tree() {
    assert_eq!(
        run_js(
            r#"
function hasPathSum(root, target) {
    if (!root) return false;
    if (!root.left && !root.right) return root.val === target;
    return hasPathSum(root.left, target - root.val) || hasPathSum(root.right, target - root.val);
}
const t = { val: 5, left: { val: 4, left: { val: 11, left: { val: 7, left:null, right:null }, right: { val: 2, left:null, right:null } }, right: null }, right: { val: 8, left:null, right:null } };
console.log(hasPathSum(t, 22));
console.log(hasPathSum(t, 10));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn detect_cycle() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn level_order_bfs() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["1", "2,3", "4"]
    );
}

#[test]
fn dijkstra_shortest_path() {
    assert_eq!(
        run_js(
            r#"
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
"#
        ),
        vec!["0", "1", "3", "4"]
    );
}

#[test]
fn flatten_tree_to_list() {
    assert_eq!(
        run_js(
            r#"
function flatten(root, acc=[]) {
    if (!root) return acc;
    acc.push(root.val);
    flatten(root.left, acc);
    flatten(root.right, acc);
    return acc;
}
const t = { val:1, left:{val:2,left:{val:4,left:null,right:null},right:{val:5,left:null,right:null}}, right:{val:3,left:null,right:null} };
console.log(flatten(t).join(","));
"#
        ),
        vec!["1,2,4,5,3"]
    );
}
