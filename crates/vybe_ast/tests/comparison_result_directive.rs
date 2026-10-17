use vybe_ast::Directives;

#[test]
fn comparison_result_policy_is_unspecified_by_default() {
    assert_eq!(Directives::default().comparison_results_are_boolean, None);
}

#[test]
fn unspecified_child_inherits_boolean_comparison_values() {
    let mut inherited = Directives {
        comparison_results_are_boolean: Some(true),
        ..Default::default()
    };
    inherited.overlay(&Directives::default());
    assert_eq!(inherited.comparison_results_are_boolean, Some(true));
}

#[test]
fn integer_comparison_values_can_override_boolean_parent() {
    let mut inherited = Directives {
        comparison_results_are_boolean: Some(true),
        ..Default::default()
    };
    inherited.overlay(&Directives {
        comparison_results_are_boolean: Some(false),
        ..Default::default()
    });
    assert_eq!(inherited.comparison_results_are_boolean, Some(false));
}
