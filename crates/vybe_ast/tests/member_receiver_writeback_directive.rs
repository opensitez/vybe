use vybe_ast::Directives;

#[test]
fn receiver_writeback_default_inheritance_and_explicit_disable() {
    let root = Directives::default();
    assert!(!root.member_receiver_writeback());
    assert!(root.is_empty());
    let mut enabled = root.clone();
    enabled.overlay(&Directives {
        member_receiver_writeback: Some(true),
        ..Default::default()
    });
    assert!(enabled.member_receiver_writeback());
    enabled.overlay(&Directives::default());
    assert!(enabled.member_receiver_writeback(), "None must inherit");
    let mut disabled = enabled.clone();
    disabled.overlay(&Directives {
        member_receiver_writeback: Some(false),
        ..Default::default()
    });
    assert!(!disabled.member_receiver_writeback());
    assert!(!disabled.is_empty(), "explicit false is a policy delta");
    assert!(enabled.member_receiver_writeback(), "popping restores parent");
    disabled.overlay(&Directives {
        member_receiver_writeback: Some(true),
        ..Default::default()
    });
    assert!(disabled.member_receiver_writeback());
}
