! vybe-test: fortran/types/derived_type_with_methods_runtime
! origin: languages/fortran/tests/fortran/test_types.rs
module m
    type :: Counter
        integer :: value = 0
    contains
        procedure :: increment
    end type Counter
contains
    subroutine increment(self)
        class(Counter), intent(inout) :: self
        self%value = self%value + 1
    end subroutine increment
end module m
program driver
use m
    type(Counter) :: c
    call c%increment()
    call c%increment()
    if ((c%value) /= 2) then
    print *, "FAIL: want [2] got [", c%value, "]"
    stop 1
end if
end program driver
