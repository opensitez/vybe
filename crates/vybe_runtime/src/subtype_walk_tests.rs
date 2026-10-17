use super::{TypeDef, TypeRegistry};
use std::{collections::HashSet, hint::black_box, time::Instant};

fn previous(reg: &TypeRegistry, id: usize, target: usize) -> bool {
    if id == target { return true; }
    let Some(root) = reg.types.get(id) else { return false; };
    if root.parent == Some(target) || root.implements.contains(&target) { return true; }
    if root.parent.is_none() && root.implements.is_empty() { return false; }
    let mut pending = root.implements.clone();
    if let Some(parent) = root.parent { pending.push(parent); }
    let mut seen = HashSet::new();
    seen.insert(id);
    while let Some(tid) = pending.pop() {
        if tid == target { return true; }
        if !seen.insert(tid) { continue; }
        if let Some(ty) = reg.types.get(tid) {
            pending.extend(ty.implements.iter().copied());
            if let Some(parent) = ty.parent { pending.push(parent); }
        }
    }
    false
}

#[test]
fn parent_fast_path_preserves_graph_traversal_results() {
    for encoding in 0..6usize.pow(4) {
        let mut graph = TypeRegistry::new();
        for index in 1..4 { graph.register(TypeDef::new(&format!("type_{index}"))); }
        let mut code = encoding;
        for index in 0..4 {
            let parent = code % 6;
            code /= 6;
            graph.types[index].parent = (parent != 5).then_some(parent);
        }
        for interfaces in 0..16usize {
            for index in 0..4 {
                graph.types[index].implements = if interfaces & (1 << index) != 0 {
                    vec![(index + 2) % 4, 4]
                } else { Vec::new() };
            }
            for id in 0..6 { for target in 0..6 {
                assert_eq!(graph.is_subtype(id, target), previous(&graph, id, target),
                    "graph={encoding} interfaces={interfaces} id={id} target={target}");
            } }
        }
    }
}

fn timed(f: fn(&TypeRegistry, usize, usize) -> bool, reg: &TypeRegistry,
    id: usize, target: usize) -> f64
{
    let start = Instant::now();
    for _ in 0..20000 {
        black_box(f(black_box(reg), black_box(id), black_box(target)));
    }
    start.elapsed().as_secs_f64() * 1e9 / 20000.0
}

#[test]
fn interface_suffix_traversal_preserves_long_cycles_and_invalid_edges() {
    for depth in [8usize, 32, 64] {
        for interface_at in [0, 1, depth / 2, depth - 1, depth] {
            for edge in [0, interface_at, depth, depth + 1, usize::MAX] {
                let mut reg = TypeRegistry::new();
                for index in 1..=depth {
                    reg.register(TypeDef::new(&format!("type_{index}")).with_parent(index - 1));
                }
                reg.types[interface_at].implements.extend([edge, interface_at]);
                // The graph may re-enter the prefix after entering interfaces.
                for target in (0..=depth + 1).chain(std::iter::once(usize::MAX)) {
                    assert_eq!(reg.is_subtype(depth, target), previous(&reg, depth, target),
                        "depth={depth} interface_at={interface_at} edge={edge} target={target}");
                }
                reg.types[0].parent = Some(depth);
                for target in (0..=depth + 1).chain(std::iter::once(usize::MAX)) {
                    assert_eq!(reg.is_subtype(depth, target), previous(&reg, depth, target),
                        "cyclic depth={depth} interface_at={interface_at} edge={edge} target={target}");
                }
            }
        }
    }
}

#[test]
fn branching_graphs_preserve_results_across_inline_visited_limit() {
    for width in [0usize, 1, 7, 8, 9, 32] {
        let mut reg = TypeRegistry::new();
        for id in 1..=width + 1 { reg.register(TypeDef::new(&format!("branch_{id}"))); }
        reg.types[width + 1].implements.extend(0..width);
        for id in 0..width {
            reg.types[id].parent = Some(width + 1);
            reg.types[id].implements.extend([id, width + 2]);
        }
        for id in 0..=width + 2 {
            for target in (0..=width + 2).chain(std::iter::once(usize::MAX)) {
                assert_eq!(reg.is_subtype(id, target), previous(&reg, id, target),
                    "width={width} id={id} target={target}");
            }
        }
    }
}

#[test]
#[ignore = "paired native RTT traversal timing"]
fn parent_chain_subtype_benchmark() {
    for depth in [1usize, 2, 4, 8, 32] {
        for interfaces in [false, true] {
            let mut reg = TypeRegistry::new();
            for index in 1..=depth {
                reg.register(TypeDef::new(&format!("type_{index}")).with_parent(index - 1));
            }
            if interfaces { reg.types[depth / 2].implements.push(depth + 2); }
            for target in [0, depth + 1, depth + 2] {
                let mut old = Vec::new();
                let mut new = Vec::new();
                for trial in 0..15 {
                    if trial % 2 == 0 {
                        old.push(timed(previous, &reg, depth, target));
                        new.push(timed(TypeRegistry::is_subtype, &reg, depth, target));
                    } else {
                        new.push(timed(TypeRegistry::is_subtype, &reg, depth, target));
                        old.push(timed(previous, &reg, depth, target));
                    }
                }
                old.sort_by(f64::total_cmp);
                new.sort_by(f64::total_cmp);
                println!("depth={depth} interfaces={interfaces} target={target} previous_ns={:.2} current_ns={:.2} ratio={:.2}",
                    old[7], new[7], old[7] / new[7]);
            }
        }
    }
}
