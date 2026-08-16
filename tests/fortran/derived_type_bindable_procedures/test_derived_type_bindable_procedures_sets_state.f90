! vybe-test: fortran/derived_type_bindable_procedures/test_derived_type_bindable_procedures_sets_state
! origin: languages/fortran/tests/fortran/test_derived_type_bindable_procedures.rs
module m
    type :: counter
        integer :: value = 1
    contains
        procedure :: inc => counter_inc
    end type
contains
    subroutine counter_inc(self, delta)
        class(counter), intent(inout) :: self
        integer, intent(in) :: delta
        self%value = self%value + delta
    end subroutine
end module m
program driver
use m

    type(counter) :: c
    call c%inc(4)
    if ((c%value) /= 5) then
    print *, "FAIL: want [5] got [", c%value, "]"
    stop 1
end if

end program driver
