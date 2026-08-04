! vybe-test: fortran/types/derived_type_with_methods
! origin: languages/fortran/tests/fortran/test_types.rs

program test
    type :: Counter
        integer :: value = 0
    contains
        procedure :: increment
    end type Counter
    type(Counter) :: c
    if ((c%value) /= 1) then
    print *, "FAIL: want [1] got [", c%value, "]"
    stop 1
end if
contains
    subroutine increment(self)
        class(Counter), intent(inout) :: self
        self%value = self%value + 1
    end subroutine increment
end program test
