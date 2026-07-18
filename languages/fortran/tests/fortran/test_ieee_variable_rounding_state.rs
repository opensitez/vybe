use super::helpers::run_prints;

#[test]
fn test_ieee_variable_rounding_state_reads_current_mode() {
    let out = run_prints(
        r#"
program test_ieee_variable_rounding_state
    use, intrinsic :: ieee_arithmetic
    integer :: mode
    call ieee_get_rounding_mode(mode)
    print *, mode
end program test_ieee_variable_rounding_state
"#,
    );

    assert_eq!(out.len(), 1);
}
