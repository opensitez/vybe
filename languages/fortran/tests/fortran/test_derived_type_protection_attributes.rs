use super::helpers::run_prints;

#[test]
fn test_derived_type_protection_attributes_private_public_members() {
    let out = run_prints(
        r#"
module derived_type_protection_attributes_module
    type :: item
        integer :: public_value
    contains
        procedure :: value_copy
    end type

    function value_copy(self) result(n)
        class(item), intent(in) :: self
        integer :: n
        n = self%public_value
    end function
end module

program test_derived_type_protection_attributes
    use derived_type_protection_attributes_module
    type(item) :: token
    token%public_value = 8
    print *, value_copy(token)
end program test_derived_type_protection_attributes
"#,
    );

    assert_eq!(out, vec!["8"]);
}
