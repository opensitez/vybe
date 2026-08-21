! vybe-test: fortran/type_bound_procedures/type_bound_generic_binding_alias_runs
! origin: languages/fortran/tests/fortran/test_fortran2003.rs

program test
    type :: Counter
        integer :: n = 4
    contains
        procedure :: doubled_impl
        generic :: doubled => doubled_impl
    end type Counter

    type(Counter) :: value
    if ((value%doubled()) /= 8) then
    print *, "FAIL: want [8] got [", value%doubled(), "]"
    stop 1
end if
contains
    integer function doubled_impl(self) result(v)
        class(Counter), intent(in) :: self
        integer :: v
        v = self%n * 2
    end function doubled_impl
end program test
