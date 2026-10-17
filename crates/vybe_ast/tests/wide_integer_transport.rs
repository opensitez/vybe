use vybe_ast::{
    DirectiveScope, Directives, ShiftOverflow, Statement, StmtKind, WideIntegerTransport,
};

fn delta(transport: WideIntegerTransport) -> Directives {
    Directives {
        wide_integer_transport: Some(transport),
        ..Default::default()
    }
}

// Model the documented directive-frame contract using the actual AST deltas.
// Compiler emission/region integration is a separate gate once consumers land.
fn apply(frames: &mut [Directives], statement: &Statement) {
    let StmtKind::Directive { set, scope } = &statement.kind else {
        panic!("expected directive delta");
    };
    match scope {
        DirectiveScope::Block => frames.last_mut().unwrap().overlay(set),
        DirectiveScope::Module => {
            for frame in frames {
                frame.overlay(set);
            }
        }
    }
}

#[test]
fn unstated_transport_is_empty_and_effectively_number() {
    let directives = Directives::default();
    assert_eq!(directives.wide_integer_transport, None);
    assert_eq!(
        directives.wide_integer_transport(),
        WideIntegerTransport::Number
    );
    assert_eq!(
        WideIntegerTransport::default(),
        WideIntegerTransport::Number
    );
    assert!(directives.is_empty());
    assert!(!delta(WideIntegerTransport::Number).is_empty());
    assert!(!delta(WideIntegerTransport::Exact).is_empty());
}

#[test]
fn omitted_transport_inherits_and_explicit_number_overrides_exact() {
    for transport in [WideIntegerTransport::Number, WideIntegerTransport::Exact] {
        let mut inherited = delta(transport);
        inherited.overlay(&Directives::default());
        assert_eq!(inherited.wide_integer_transport, Some(transport));
        assert_eq!(inherited.wide_integer_transport(), transport);
    }
    let mut directives = delta(WideIntegerTransport::Exact);
    directives.overlay(&delta(WideIntegerTransport::Number));
    assert_eq!(
        directives.wide_integer_transport,
        Some(WideIntegerTransport::Number)
    );
    directives.overlay(&delta(WideIntegerTransport::Exact));
    assert_eq!(
        directives.wide_integer_transport(),
        WideIntegerTransport::Exact
    );
}

#[test]
fn transport_overlay_preserves_independent_shift_policy() {
    let mut directives = Directives {
        shift_overflow: Some(ShiftOverflow::Zero),
        ..Default::default()
    };
    directives.overlay(&delta(WideIntegerTransport::Exact));
    assert_eq!(directives.shift_overflow, Some(ShiftOverflow::Zero));
    directives.overlay(&Directives {
        shift_overflow: Some(ShiftOverflow::Mask),
        ..Default::default()
    });
    assert_eq!(directives.shift_overflow, Some(ShiftOverflow::Mask));
    assert_eq!(
        directives.wide_integer_transport(),
        WideIntegerTransport::Exact
    );
}

#[test]
fn block_transport_inherits_and_restores_at_nested_region_exit() {
    let module = delta(WideIntegerTransport::Number);
    let mut frames = vec![module.clone(), module.clone()];
    apply(
        &mut frames,
        &Statement::new(StmtKind::Directive {
            set: delta(WideIntegerTransport::Exact),
            scope: DirectiveScope::Block,
        }),
    );
    assert_eq!(frames[0], module);
    assert_eq!(
        frames.last().unwrap().wide_integer_transport(),
        WideIntegerTransport::Exact
    );
    frames.push(frames.last().unwrap().clone());
    apply(
        &mut frames,
        &Statement::new(StmtKind::Directive {
            set: delta(WideIntegerTransport::Number),
            scope: DirectiveScope::Block,
        }),
    );
    assert_eq!(
        frames.last().unwrap().wide_integer_transport(),
        WideIntegerTransport::Number
    );
    frames.pop();
    assert_eq!(
        frames.last().unwrap().wide_integer_transport(),
        WideIntegerTransport::Exact
    );
    frames.pop();
    assert_eq!(frames[0], module);
}

#[test]
fn module_transport_updates_all_frames_and_survives_region_exit() {
    for transport in [WideIntegerTransport::Number, WideIntegerTransport::Exact] {
        let initial = if transport == WideIntegerTransport::Number {
            WideIntegerTransport::Exact
        } else {
            WideIntegerTransport::Number
        };
        let mut frames = vec![delta(initial), delta(initial), delta(initial)];
        apply(
            &mut frames,
            &Statement::new(StmtKind::Directive {
                set: delta(transport),
                scope: DirectiveScope::Module,
            }),
        );
        assert!(
            frames
                .iter()
                .all(|frame| frame.wide_integer_transport() == transport)
        );
        frames.pop();
        frames.pop();
        assert_eq!(frames[0].wide_integer_transport, Some(transport));
    }
}
