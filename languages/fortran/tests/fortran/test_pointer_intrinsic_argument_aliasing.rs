use super::helpers::run_prints;

#[test]
fn test_pointer_intrinsic_argument_aliasing_with_assumed_shape() {
    let out = run_prints(
        r#"
program test_pointer_intrinsic_argument_aliasing
    integer, target :: storage
    integer, pointer :: p
    storage = 4
    p => storage
    call mutate(p)
    print *, storage

contains
    subroutine mutate(value)
        integer, pointer, intent(inout) :: value
        value = 11
    end subroutine
end program test_pointer_intrinsic_argument_aliasing
"#,
    );

    assert_eq!(out, vec!["11"]);
}
