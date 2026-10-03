use vybe_parser::arena::Arena;

// Intentionally not Clone: rollback must retain handles, not clone AST nodes.
#[derive(Debug)]
enum Node {
    Number(i64),
    Add(vybe_parser::arena::Id, vybe_parser::arena::Id),
}

#[test]
fn rollback_restores_popped_values_and_discards_speculative_nodes() {
    let mut arena = Arena::default();
    let one = arena.alloc(Node::Number(1));
    arena.push(one);
    let mark = arena.checkpoint();
    let two = arena.alloc(Node::Number(2));
    arena.push(two);
    let right = arena.pop().unwrap();
    let left = arena.pop().unwrap();
    let sum = arena.alloc(Node::Add(left, right));
    arena.push(sum);
    assert!(matches!(arena.get(sum), Node::Add(a, b) if *a == one && *b == two));
    arena.rollback(mark);
    assert_eq!(arena.values(), &[one]);
    assert_eq!(arena.node_count(), 1);
    assert!(matches!(arena.get(one), Node::Number(1)));
    arena.rollback(mark); // restoring the same still-valid mark is idempotent
}

#[test]
fn nested_checkpoints_and_commit_keep_committed_ids() {
    let mut arena = Arena::default();
    let one = arena.alloc(Node::Number(1));
    arena.push(one);
    arena.commit();
    let outer = arena.checkpoint();
    let two = arena.alloc(Node::Number(2));
    arena.push(two);
    let inner = arena.checkpoint();
    assert_eq!(arena.pop(), Some(two));
    assert_eq!(arena.pop(), Some(one));
    arena.rollback(inner);
    assert_eq!(arena.values(), &[one, two]);
    arena.rollback(outer);
    assert_eq!(arena.values(), &[one]);
    assert_eq!(arena.node_count(), 1);
}
