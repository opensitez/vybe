use super::helpers::run_prints;

#[test]
fn test_execution_order_for_pure_functions_is_deterministic() {
    let out = run_prints(
        r#"
program test_execution_order_for_pure_functions
    integer :: value
    value = one_plus_two() + one_plus_two()
    print *, value

contains
    pure function one_plus_two() result(r)
        integer :: r
        r = 1 + 2
    end function
end program test_execution_order_for_pure_functions
"#,
    );

    assert_eq!(out, vec!["6"]);
}
