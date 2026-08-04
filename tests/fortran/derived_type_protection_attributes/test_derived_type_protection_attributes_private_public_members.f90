! vybe-test: fortran/derived_type_protection_attributes/test_derived_type_protection_attributes_private_public_members
! origin: languages/fortran/tests/fortran/test_derived_type_protection_attributes.rs

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
    if ((value_copy(token)) /= 8) then
    print *, "FAIL: want [8] got [", value_copy(token), "]"
    stop 1
end if
end program test_derived_type_protection_attributes
