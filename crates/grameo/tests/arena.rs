use grameo::arena::Arena;

// Intentionally not Clone: rollback must retain handles, not clone AST nodes.
#[derive(Debug)]
enum Node {
    Number(i64),
    Add(grameo::arena::Id, grameo::arena::Id),
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

#[test]
fn owned_nonclone_children_move_into_parents_and_are_extracted_on_rollback() {
    use grameo::arena::OwnedArena;
    #[derive(Debug, PartialEq)]
    enum Owned {
        Number(i64),
        Neg(Box<Owned>),
        Add(Box<Owned>, Box<Owned>),
    }
    fn split_neg(parent: Owned) -> Owned {
        match parent {
            Owned::Neg(child) => *child,
            _ => panic!("wrong inverse"),
        }
    }
    fn split_add(parent: Owned) -> (Owned, Owned) {
        match parent {
            Owned::Add(a, b) => (*a, *b),
            _ => panic!("wrong inverse"),
        }
    }
    let mut arena = OwnedArena::default();
    let one = arena.alloc(Owned::Number(1));
    arena.push(one);
    let mark = arena.checkpoint();
    let two = arena.alloc(Owned::Number(2));
    arena.push(two);
    let right = arena.pop().unwrap();
    let left = arena.pop().unwrap();
    let sum = arena.binary(
        left,
        right,
        |a, b| Owned::Add(Box::new(a), Box::new(b)),
        split_add,
    );
    arena.push(sum);
    let inner = arena.checkpoint();
    let sum = arena.pop().unwrap();
    let neg = arena.unary(sum, |a| Owned::Neg(Box::new(a)), split_neg);
    arena.push(neg);
    arena.rollback(inner);
    assert_eq!(arena.values(), &[sum]);
    assert_eq!(
        arena.get(sum),
        &Owned::Add(Box::new(Owned::Number(1)), Box::new(Owned::Number(2)))
    );
    arena.rollback(mark);
    arena.rollback(mark);
    assert_eq!(arena.values(), &[one]);
    assert_eq!(arena.get(one), &Owned::Number(1));
    let one = arena.pop().unwrap();
    let neg = arena.unary(one, |a| Owned::Neg(Box::new(a)), split_neg);
    arena.push(neg);
    arena.commit();
    assert_eq!(arena.into_root(neg), Owned::Neg(Box::new(Owned::Number(1))));
}

#[test]
fn metadata_edits_survive_moves_and_restore_exact_nonclone_children() {
    use grameo::arena::MetadataArena;
    #[derive(Debug, PartialEq)]
    struct Located {
        span: (usize, usize),
        child: Option<Box<Located>>,
    }
    fn restore(value: &mut Located, previous: (usize, usize)) {
        value.span = previous;
    }
    fn split(value: Located) -> Located {
        *value.child.unwrap()
    }
    let mut arena = MetadataArena::default();
    let child = arena.alloc(Located {
        span: (1, 2),
        child: None,
    });
    arena.push(child);
    arena.commit();
    let outer = arena.checkpoint();
    arena.edit(
        child,
        |value| std::mem::replace(&mut value.span, (0, 3)),
        restore,
    );
    let inner = arena.checkpoint();
    let child = arena.pop().unwrap();
    let parent = arena.unary(
        child,
        |child| Located {
            span: (0, 5),
            child: Some(Box::new(child)),
        },
        split,
    );
    arena.push(parent);
    arena.edit(
        parent,
        |value| std::mem::replace(&mut value.span, (0, 7)),
        restore,
    );
    assert_eq!(arena.get(parent).child.as_ref().unwrap().span, (0, 3));
    arena.rollback(inner);
    assert_eq!(arena.values(), &[child]);
    assert_eq!(arena.get(child).span, (0, 3));
    arena.edit(
        child,
        |value| std::mem::replace(&mut value.span, (9, 10)),
        restore,
    );
    arena.rollback(outer);
    assert_eq!(
        arena.get(child),
        &Located {
            span: (1, 2),
            child: None
        }
    );
    assert_eq!(arena.values(), &[child]);
    arena.edit(
        child,
        |value| std::mem::replace(&mut value.span, (0, 3)),
        restore,
    );
    arena.commit();
    assert_eq!(arena.into_root(child).span, (0, 3));
}

#[test]
fn ordered_owned_lists_restore_nested_nonclone_children_and_metadata() {
    use grameo::arena::MetadataArena;
    #[derive(Debug, PartialEq)]
    enum Node {
        Number(usize),
        List(Vec<Node>),
    }
    fn split(node: Node) -> Vec<Node> {
        let Node::List(values) = node else {
            panic!("wrong inverse")
        };
        values
    }
    let mut arena = MetadataArena::default();
    let children: Vec<_> = (0..64)
        .map(|value| arena.alloc(Node::Number(value)))
        .collect();
    let mark = arena.checkpoint();
    arena.edit(
        children[0],
        |node| {
            let Node::Number(value) = node else {
                unreachable!()
            };
            std::mem::replace(value, 100)
        },
        |node, previous| *node = Node::Number(previous),
    );
    let inner = arena.checkpoint();
    let mut order = children.clone();
    order.reverse();
    let list = arena.list(order, Node::List, split);
    let empty = arena.list(Vec::new(), Node::List, split);
    let outer = arena.list(vec![list, empty], Node::List, split);
    let Node::List(values) = arena.get(outer) else {
        panic!("expected list")
    };
    assert_eq!(values[1], Node::List(Vec::new()));
    let Node::List(values) = &values[0] else {
        panic!("expected nested list")
    };
    assert_eq!(values[0], Node::Number(63));
    assert_eq!(values[63], Node::Number(100));
    arena.rollback(inner);
    assert_eq!(arena.get(children[0]), &Node::Number(100));
    arena.rollback(mark);
    for (value, child) in children.iter().enumerate() {
        assert_eq!(arena.get(*child), &Node::Number(value));
    }
    let list = arena.list(children, Node::List, split);
    arena.commit();
    let Node::List(values) = arena.into_root(list) else {
        panic!("expected owned list")
    };
    assert_eq!(values.len(), 64);
}

#[test]
fn aliased_list_children_are_rejected_before_any_ownership_changes() {
    use grameo::arena::OwnedArena;
    let mut arena = OwnedArena::default();
    let child = arena.alloc(String::from("unchanged"));
    let mark = arena.checkpoint();
    for count in [2, 32] {
        let error = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            arena.list(
                vec![child; count],
                |values| values.concat(),
                |value| vec![value],
            );
        }));
        assert!(error.is_err());
        assert_eq!(arena.checkpoint(), mark);
        assert_eq!(arena.get(child), "unchanged");
    }
}

#[test]
fn flat_binary_metadata_restores_exact_owned_children_and_commits_multiple_roots() {
    use grameo::arena::MetadataArena;
    struct Value { text: String, offset: usize }
    let mut arena = MetadataArena::<Value, (usize, usize)>::default();
    let left = arena.alloc(Value { text: "left".into(), offset: 7 });
    let right = arena.alloc(Value { text: "right".into(), offset: 31 });
    let mark = arena.checkpoint();
    let parent = arena.binary_with_metadata(left, right, |left, right| {
        (Value { text: format!("{}|{}", left.text, right.text), offset: 0 }, (left.offset, right.offset))
    }, |value, (left, right)| {
        let (a, b) = value.text.split_once('|').unwrap();
        (Value { text: a.into(), offset: left }, Value { text: b.into(), offset: right })
    });
    assert_eq!(arena.get(parent).text, "left|right");
    arena.rollback(mark);
    assert_eq!((&*arena.get(left).text, arena.get(left).offset), ("left", 7));
    assert_eq!((&*arena.get(right).text, arena.get(right).offset), ("right", 31));
    let values: Vec<_> = arena.into_live_nodes().map(|v| (v.text, v.offset)).collect();
    assert_eq!(values, [("left".into(), 7), ("right".into(), 31)]);
}

#[test]
fn list_lowering_inverse_restores_owned_children_after_nested_speculation() {
    use grameo::arena::MetadataArena;
    // Deliberately non-Clone. Lowering changes the values and discards their
    // source offsets; the journal retains only the offsets needed for reversal.
    struct Value { values: Vec<i64>, offsets: Vec<usize> }
    let mut arena = MetadataArena::<Value, Vec<usize>>::default();
    let children: Vec<_> = [(3, 7), (5, 19), (8, 31)].into_iter()
        .map(|(n, offset)| arena.alloc(Value { values: vec![n], offsets: vec![offset] }))
        .collect();
    let mark = arena.checkpoint();
    let parent = arena.list_with_metadata(children.clone(), |children| {
        let offsets = children.iter().map(|v| v.offsets[0]).collect();
        (Value { values: children.into_iter().map(|v| v.values[0] + 10).collect(), offsets: vec![] }, offsets)
    }, |value, offsets| {
        value.values.into_iter().zip(offsets)
            .map(|(n, offset)| Value { values: vec![n - 10], offsets: vec![offset] }).collect()
    });
    let nested = arena.checkpoint();
    let wrapper = arena.list_with_metadata(vec![parent], |values| {
        (values.into_iter().next().unwrap(), vec![])
    }, |value, _| vec![value]);
    assert_eq!(arena.get(wrapper).values, [13, 15, 18]);
    arena.rollback(nested);
    assert_eq!(arena.get(parent).values, [13, 15, 18]);
    arena.rollback(mark);
    for (id, (n, offset)) in children.into_iter().zip([(3, 7), (5, 19), (8, 31)]) {
        assert_eq!(arena.get(id).values, [n]);
        assert_eq!(arena.get(id).offsets, [offset]);
    }
    arena.rollback(mark);
}

#[test]
fn separate_edit_payload_restores_interleaved_edits_and_compositions() {
    use grameo::arena::EditArena;
    // Distinct payload sizes share one cursor. The values are deliberately
    // non-Clone; composing and undoing must preserve exclusive ownership.
    enum Value { Leaf(Box<u8>, u32), Pair(Box<Value>, Box<Value>, u32) }
    fn edit(value: &mut Value, span: u32) -> u32 {
        let (Value::Leaf(_, current) | Value::Pair(_, _, current)) = value;
        std::mem::replace(current, span)
    }
    fn restore(value: &mut Value, previous: u32) { edit(value, previous); }
    let mut arena = EditArena::<Value, [u8; 128], u32>::default();
    let left = arena.alloc(Value::Leaf(Box::new(3), 10));
    let right = arena.alloc(Value::Leaf(Box::new(7), 20));
    arena.push(left);
    arena.push(right);
    let original = arena.checkpoint();
    arena.edit(left, |value| edit(value, 11), restore);
    let edited = arena.checkpoint();
    arena.pop();
    arena.pop();
    let parent = arena.binary_with_metadata(left, right, |left, right| {
        (Value::Pair(Box::new(left), Box::new(right), 0), [0; 128])
    }, |value, _| {
        let Value::Pair(left, right, _) = value else { panic!("pair") };
        (*left, *right)
    });
    arena.push(parent);
    let composed = arena.checkpoint();
    arena.edit(parent, |value| edit(value, 99), restore);
    arena.rollback(composed);
    assert!(matches!(arena.get(parent), Value::Pair(_, _, 0)));
    arena.rollback(edited);
    assert!(matches!(arena.get(left), Value::Leaf(value, 11) if **value == 3));
    assert_eq!(arena.values(), &[left, right]);
    arena.rollback(original);
    assert!(matches!(arena.get(left), Value::Leaf(value, 10) if **value == 3));
    assert!(matches!(arena.get(right), Value::Leaf(value, 20) if **value == 7));
    arena.rollback(original);
}

#[derive(Debug)]
struct PrefixNode {
    kind: PrefixKind,
    span: (usize, usize),
}
#[derive(Debug)]
enum PrefixKind {
    Leaf(String),
    Prefix(Box<PrefixNode>),
}
fn wrap_prefix(mut child: PrefixNode) -> (PrefixNode, (usize, usize)) {
    let previous = child.span;
    child.span = (0, 0);
    (PrefixNode { kind: PrefixKind::Prefix(Box::new(child)), span: (previous.0 - 1, previous.1) }, previous)
}
fn unwrap_prefix(parent: PrefixNode, previous: (usize, usize)) -> PrefixNode {
    let PrefixKind::Prefix(child) = parent.kind else { panic!("prefix inverse"); };
    let mut child = *child;
    child.span = previous;
    child
}
#[test]
fn metadata_prefix_rollback_restores_nested_owned_operands_and_stack() {
    let mut arena = grameo::arena::MetadataArena::<PrefixNode, (usize, usize)>::default();
    let leaf = arena.alloc(PrefixNode { kind: PrefixKind::Leaf("operand".into()), span: (64, 72) });
    arena.push(leaf);
    let mut states = Vec::new();
    for _ in 0..32 {
        let child = *arena.values().last().unwrap();
        states.push((arena.checkpoint(), child, format!("{:?}", arena.get(child))));
        assert_eq!(arena.pop(), Some(child));
        let parent = arena.unary_with_metadata(child, wrap_prefix, unwrap_prefix);
        arena.push(parent);
    }
    for (mark, child, expected) in states.into_iter().rev() {
        arena.rollback(mark);
        assert_eq!(arena.values(), &[child]);
        assert_eq!(format!("{:?}", arena.get(child)), expected);
    }
    let PrefixKind::Leaf(name) = &arena.get(leaf).kind else { panic!("leaf restored"); };
    assert_eq!(name, "operand");
}
#[test]
fn metadata_prefix_commit_keeps_owned_tree_for_later_speculation() {
    let mut arena = grameo::arena::MetadataArena::<PrefixNode, (usize, usize)>::default();
    let leaf = arena.alloc(PrefixNode { kind: PrefixKind::Leaf("operand".into()), span: (64, 72) });
    let committed = arena.unary_with_metadata(leaf, wrap_prefix, unwrap_prefix);
    arena.push(committed);
    arena.commit();
    let before = format!("{:?}", arena.get(committed));
    let mark = arena.checkpoint();
    arena.pop();
    let speculative = arena.unary_with_metadata(committed, wrap_prefix, unwrap_prefix);
    arena.push(speculative);
    arena.rollback(mark);
    assert_eq!(arena.values(), &[committed]);
    assert_eq!(format!("{:?}", arena.get(committed)), before);
}
