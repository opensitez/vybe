use super::helpers::run_prints;

#[test]
fn test_derived_type_intrinsics_interop_bind_c_shape_query() {
    let out = run_prints(
        r#"
program test_derived_type_intrinsics_interop
    use iso_c_binding, only: c_int
    type, bind(C) :: payload
        integer(c_int) :: value
    end type

    type(payload) :: p
    p%value = 9
    print *, p%value
end program test_derived_type_intrinsics_interop
"#,
    );

    assert_eq!(out, vec!["9"]);
}
