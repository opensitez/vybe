! vybe-test: fortran/derived_types_advanced/type_bound_final
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

program test
    type :: Resource
        integer :: id
    contains
        final :: cleanup
    end type Resource
contains
    subroutine cleanup(self)
        type(Resource), intent(inout) :: self
        self%id = 0
    end subroutine cleanup
end program test
