! vybe-test: fortran/derived_types_advanced/module_type_bound_subroutine
! origin: languages/fortran/tests/fortran/test_derived_types_advanced.rs

module counters
    implicit none

    type :: Counter
        integer :: n = 0
    contains
        procedure :: inc
        procedure :: get
    end type Counter
contains
    subroutine inc(self)
        class(Counter), intent(inout) :: self
        self%n = self%n + 1
    end subroutine inc

    function get(self) result(v)
        class(Counter), intent(in) :: self
        integer :: v
        v = self%n
    end function get
end module counters

program test
    use counters
    implicit none
    type(Counter) :: c
    call c%inc()
    call c%inc()
    if ((c%get()) /= 2) then
    print *, "FAIL: want [2] got [", c%get(), "]"
    stop 1
end if
end program test
