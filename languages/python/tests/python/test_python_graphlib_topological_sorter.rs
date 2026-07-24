use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Category 2: Graph Dependency Resolution (graphlib module)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_graphlib_static_order_basic() {
    let out = run_python(r#"
import graphlib

graph = {"D": {"B", "C"}, "C": {"A"}, "B": {"A"}}
ts = graphlib.TopologicalSorter(graph)
order = list(ts.static_order())
print(order[0] == "A")
print(order[-1] == "D")
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_graphlib_add_method() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter()
ts.add("cook", "prep")
ts.add("eat", "cook")
order = list(ts.static_order())
print(order)
"#);
    assert_eq!(out, vec!["['prep', 'cook', 'eat']"]);
}

#[test]
fn test_graphlib_cycle_detection() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter()
ts.add("A", "B")
ts.add("B", "C")
ts.add("C", "A")

try:
    list(ts.static_order())
except graphlib.CycleError as e:
    print("CycleErrorCaught")
"#);
    assert_eq!(out, vec!["CycleErrorCaught"]);
}

#[test]
fn test_graphlib_step_by_step_prepare_get_ready() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter()
ts.add("B", "A")
ts.add("C", "B")
ts.prepare()

ready1 = ts.get_ready()
print(ready1)
ts.done("A")
ready2 = ts.get_ready()
print(ready2)
"#);
    assert_eq!(out, vec!["('A',)", "('B',)"]);
}

#[test]
fn test_graphlib_is_active() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"B": {"A"}})
ts.prepare()
print(ts.is_active())
nodes = ts.get_ready()
ts.done(*nodes)
nodes2 = ts.get_ready()
ts.done(*nodes2)
print(ts.is_active())
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_graphlib_add_multiple_predecessors() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter()
ts.add("job3", "job1", "job2")
order = list(ts.static_order())
print(order[-1])
"#);
    assert_eq!(out, vec!["job3"]);
}

#[test]
fn test_graphlib_empty_graph() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({})
print(list(ts.static_order()))
"#);
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_graphlib_single_node() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"A": set()})
print(list(ts.static_order()))
"#);
    assert_eq!(out, vec!["['A']"]);
}

#[test]
fn test_graphlib_prepare_twice_raises_value_error() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"A": set()})
ts.prepare()
try:
    ts.prepare()
except ValueError:
    print("ValueErrorCaught")
"#);
    assert_eq!(out, vec!["ValueErrorCaught"]);
}

#[test]
fn test_graphlib_add_after_prepare_raises_value_error() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"A": set()})
ts.prepare()
try:
    ts.add("B", "A")
except ValueError:
    print("ValueErrorCaught")
"#);
    assert_eq!(out, vec!["ValueErrorCaught"]);
}

#[test]
fn test_graphlib_done_unready_node_raises_value_error() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"B": {"A"}})
ts.prepare()
try:
    ts.done("B") # B is not ready yet!
except ValueError:
    print("ValueErrorCaught")
"#);
    assert_eq!(out, vec!["ValueErrorCaught"]);
}

#[test]
fn test_graphlib_linear_chain() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"3": {"2"}, "2": {"1"}})
print(list(ts.static_order()))
"#);
    assert_eq!(out, vec!["['1', '2', '3']"]);
}

#[test]
fn test_graphlib_diamond_dependency() {
    let out = run_python(r#"
import graphlib

graph = {"D": {"B", "C"}, "B": {"A"}, "C": {"A"}}
ts = graphlib.TopologicalSorter(graph)
order = list(ts.static_order())
print(len(order))
print(order[0])
print(order[-1])
"#);
    assert_eq!(out, vec!["4", "A", "D"]);
}

#[test]
fn test_graphlib_cycle_error_args() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"A": {"B"}, "B": {"A"}})
try:
    ts.prepare()
except graphlib.CycleError as e:
    cycle = e.args[1]
    print(len(cycle) >= 2)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_graphlib_integer_nodes() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({3: {1, 2}, 2: {1}})
print(list(ts.static_order()))
"#);
    assert_eq!(out, vec!["[1, 2, 3]"]);
}

#[test]
fn test_graphlib_disconnected_components() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"A": set(), "B": set()})
ts.prepare()
ready = ts.get_ready()
print(set(ready) == {"A", "B"})
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_graphlib_get_ready_returns_empty_when_done() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"A": set()})
ts.prepare()
nodes = ts.get_ready()
ts.done(*nodes)
print(ts.get_ready())
"#);
    assert_eq!(out, vec!["()"]);
}

#[test]
fn test_graphlib_self_loop_raises_cycle_error() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"A": {"A"}})
try:
    ts.prepare()
except graphlib.CycleError:
    print("CycleErrorCaught")
"#);
    assert_eq!(out, vec!["CycleErrorCaught"]);
}

#[test]
fn test_graphlib_static_order_returns_iterator() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter({"B": {"A"}})
it = ts.static_order()
print(next(it))
print(next(it))
"#);
    assert_eq!(out, vec!["A", "B"]);
}

#[test]
fn test_graphlib_multiple_add_calls_same_node() {
    let out = run_python(r#"
import graphlib

ts = graphlib.TopologicalSorter()
ts.add("C", "A")
ts.add("C", "B")
order = list(ts.static_order())
print(order[-1])
"#);
    assert_eq!(out, vec!["C"]);
}
