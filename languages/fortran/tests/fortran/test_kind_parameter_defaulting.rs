use super::helpers::run_prints;

#[test]
fn test_kind_parameter_defaulting_selects_default_real_kind() {
    let out = run_prints(
        r#"
program test_kind_parameter_defaulting
    integer :: k
    k = selected_real_kind(6)
    print *, k
end program test_kind_parameter_defaulting
"#,
    );

    assert_eq!(out.len(), 1);
}
