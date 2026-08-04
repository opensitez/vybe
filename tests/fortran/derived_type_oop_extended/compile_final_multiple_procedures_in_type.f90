! vybe-test: fortran/derived_type_oop_extended/compile_final_multiple_procedures_in_type
! origin: languages/fortran/tests/fortran/test_derived_type_oop_extended.rs

program t
    type :: Resource
        integer :: handle = 0
        logical :: open = .false.
    contains
        final :: close_handle
        final :: mark_closed
    end type Resource
    type(Resource) :: r
    r%handle = 42
    r%open = .true.
    print *, r%handle
contains
    subroutine close_handle(self)
        type(Resource), intent(inout) :: self
        self%handle = 0
    end subroutine close_handle
    subroutine mark_closed(self)
        type(Resource), intent(inout) :: self
        self%open = .false.
    end subroutine mark_closed
end program t
