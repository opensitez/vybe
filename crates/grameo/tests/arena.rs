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
